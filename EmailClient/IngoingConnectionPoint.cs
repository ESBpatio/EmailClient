using ESB_ConnectionPoints.PluginsInterfaces;
using Newtonsoft.Json.Linq;
using System.Threading;
using System;
using MailKit.Net.Imap;
using System.IO;
using MimeKit;
using MailKit;
using System.Text.RegularExpressions;
using ExcelDataReader;
using System.Data;
using Newtonsoft.Json;
using System.Text;
using System.Linq;
using System.Collections.Generic;

namespace EmailClient
{
    public class IngoingConnectionPoint : IStandartIngoingConnectionPoint
    {
        private readonly ILogger logger;
        private readonly IMessageFactory messageFactory;
        private string uri, login, password, classId, type;
        private bool ssl;
        private int port, timeout, startLine, sheetNumber;
        public IExcelDataReader reader;

        public IngoingConnectionPoint(string jsonSettings, IServiceLocator serviceLocator)
        {
            if (serviceLocator != null)
            {
                this.logger = serviceLocator.GetLogger(this.GetType());
                this.messageFactory = serviceLocator.GetMessageFactory();
            }
            this.ParseSettings(jsonSettings);
        }

        public void ParseSettings(string jsonSettings)
        {
            JObject jObject;
            try
            {
                jObject = JObject.Parse(jsonSettings);
            }
            catch (Exception ex)
            {

                throw new Exception("Ошибки разбора JSON настроек " + ex.Message);
            }
            this.uri = JsonUtils.StringValue(jObject, "ConfigurationServer.serverUri", "");
            this.login = JsonUtils.StringValue(jObject, "ConfigurationServer.Login", "");
            this.password = JsonUtils.StringValue(jObject, "ConfigurationServer.Password", "");
            this.ssl = JsonUtils.BoolValue(jObject, "ConfigurationServer.Ssl", false);
            this.port = JsonUtils.IntValue(jObject, "ConfigurationServer.Port", 993);
            this.type = JsonUtils.StringValue(jObject, "MessageSettings.Type", "DTP");
            this.classId = JsonUtils.StringValue(jObject, "MessageSettings.ClassId", "0");
            this.timeout = JsonUtils.IntValue(jObject, "TimeOut", 10);
        }

        public void Run(IMessageHandler messageHandler, CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                try
                {
                    GetMessage(messageHandler);
                }
                catch (Exception ex)
                {

                    logger.Error("Ошибка загрузки письма " + ex.Message);
                }
                
                ct.WaitHandle.WaitOne(timeout);
            }
        }

        public void GetMessage(IMessageHandler messageHandler)
        {
            using (ImapClient client = new ImapClient())
            {
                client.Connect(uri, port, ssl);
                client.Authenticate(login, password);
                IMailFolder inbox = client.Inbox;
                inbox.Open(FolderAccess.ReadWrite);
                for (int i = 0; i < inbox.Count; i++)
                {
                    MimeMessage emailMessage = inbox.GetMessage(i);
                    var propertyMessage = inbox.Fetch(new[] { i }, MessageSummaryItems.Flags);
                    if (!propertyMessage[0].Flags.Value.HasFlag(MessageFlags.Seen))
                    {
                        string from = emailMessage.From.OfType<MailboxAddress>().Single().Address;
                        string subject = emailMessage.Subject;
                        foreach (var attachment in emailMessage.Attachments)
                        {
                            var ext = Regex.Match(attachment.ContentType.Name, "[^.]+$").Value;
                            switch (ext)
                            {
                                case "xlsx":
                                    break;
                                case "xls":
                                    break;
                                default:
                                    logger.Warning("Файл был пропущен , формат не подходит для обработки");
                                    continue;
                            }

                            MemoryStream ms = new MemoryStream();
                            using (MemoryStream stream = ms)
                            {
                                if (attachment is MessagePart)
                                {
                                    var part = (MessagePart)attachment;
                                    part.Message.WriteTo(stream);
                                }
                                else
                                {
                                    var part = (MimePart)attachment;
                                    part.Content.DecodeTo(stream);
                                }

                                FileStream openSettings = File.Open(@"C:\excel\Settings\" + from, FileMode.Open, FileAccess.Read);
                                StreamReader sr = new StreamReader(openSettings);

                                JObject settingToProvider = JObject.Parse(sr.ReadToEnd());
                                List<rowSetting> rowSettings = GetSettingsToRows(settingToProvider);

                                if (rowSettings.Count < 0)
                                {
                                    logger.Warning(string.Format("Настройки не найдены для отправителя {0} , тема письма : {1}", from , subject));
                                    inbox.AddFlags(i, MessageFlags.Seen, true);
                                //  sr.Close();
                                    continue;
                                }    
                                //sr.Close();

                                startLine = (JsonUtils.IntValue(settingToProvider, "СхемаЗагрузки.НачальнаяСтрокаВФайле") - 1);
                                sheetNumber = (JsonUtils.IntValue(settingToProvider, "СхемаЗагрузки.НомерЛистаВФайле") - 1);

                                reader = ExcelReaderFactory.CreateReader(stream);
                                var conf = new ExcelDataSetConfiguration
                                {
                                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                                    {
                                        //UseHeaderRow = true
                                        UseHeaderRow = false,
                                        FilterRow = rowReader => rowReader.Depth > startLine
                                    }
                                };
                                DataSet dataSet = reader.AsDataSet(conf);
                                DataTable dataTable = dataSet.Tables[sheetNumber];

                                DataTable dt = new DataTable();
                                foreach (var item in rowSettings)
                                {
                                    DataColumn dataColumn = dataTable.Columns[item.numberCol - 1];
                                    DataColumn column = new DataColumn()
                                    {
                                        ColumnName = item.viewCol,
                                        DataType = dataColumn.DataType,
                                        Expression = dataColumn.Expression,
                                        ColumnMapping = dataColumn.ColumnMapping
                                    };
                                    dt.Columns.Add(column);
                                }
                                foreach (DataRow item in dataTable.Rows)
                                {
                                    DataRow newRow = dt.NewRow();
                                    for(int j = 0; j < rowSettings.Count(); j++)
                                    {
                                        newRow[dt.Columns[j].ColumnName] = item[dataTable.Columns[rowSettings[j].numberCol - 1]];
                                    }
                                    dt.Rows.Add(newRow);
                                }
                                CreateESBMessage(JsonConvert.SerializeObject(dt), messageHandler, subject, from);

                            }
                        }
                        inbox.AddFlags(i, MessageFlags.Seen, true);
                    }
                }
            }
        }


        private List<rowSetting> GetSettingsToRows(JObject jObject)
        {
            List<rowSetting> rowSettings = new List<rowSetting>();
            foreach (var item in jObject["СхемаЗагрузки"]["ЗагружаеммыеПоля"])
            {
                rowSetting rowSetting = new rowSetting
                {
                    numberCol = (int)item["НомерКолонкиВФайле"],
                    viewCol = item["ВидКолонки"].ToString()
                };
                rowSettings.Add(rowSetting);
            }
            return rowSettings;
        }

        class rowSetting
        {
            public int numberCol { get; set; }
            public string viewCol { get; set; }
        }

        public void CreateESBMessage(string jsonBody, IMessageHandler messageHandler, string subject , string from)
        {
            Message esbMessage = new Message
            {
                ClassId = classId,
                Body = Encoding.UTF8.GetBytes(jsonBody),
                Id = Guid.NewGuid(),
                Type = type,                
            };
            esbMessage.SetPropertyWithValue("from", from);
            esbMessage.SetPropertyWithValue("subject", subject);

            messageHandler.HandleMessage(esbMessage);
        }

        public void Cleanup()
        {

        }

        public void Dispose()
        {

        }

        public void Initialize()
        {

        }
    }
}
