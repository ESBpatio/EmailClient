using ESB_ConnectionPoints.PluginsInterfaces;
using ExcelDataReader;
using MailKit;
using MailKit.Net.Imap;
using MimeKit;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;


namespace EmailClient
{
    public class IngoingConnectionPoint : IStandartIngoingConnectionPoint
    {
        private readonly ILogger logger;
        private readonly IMessageFactory messageFactory;
        private string uri, login, password, classId, type, formatSetting, patchSetting, from, subject, patchToDisk, responsiblePerson, fileName, idObject;
        private bool ssl;
        private int port, timeout, startLine, sheetNumber;
        public IExcelDataReader reader;
        public EmailUtils email;
        public IMailFolder inbox;

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
            this.formatSetting = JsonUtils.StringValue(jObject, "formatSetting", ".JSON");
            this.patchSetting = JsonUtils.StringValue(jObject, "patchSetting", @"C:\Settings\");
            this.patchToDisk = JsonUtils.StringValue(jObject, "patchToDisk", @"C:\tmp\");
            this.responsiblePerson = JsonUtils.StringValue(jObject, "responsiblePerson");
        }
        public void Run(IMessageHandler messageHandler, CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                try
                {
                    GetMessages(messageHandler);
                }
                catch (Exception ex)
                {
                    string error = string.Format("Ошибка загрузки письма " + ex.Message + "\n" + ex.StackTrace);
                    logger.Error(error);
                    email.sendMessage(responsiblePerson, error, uri, 587, login, password);
                }
                ct.WaitHandle.WaitOne(timeout);
            }
        }
        public void GetMessages(IMessageHandler messageHandler)
        {
            using (ImapClient client = new ImapClient())
            {
                client.Connect(uri, port, ssl);
                client.Authenticate(login, password);
                inbox = client.Inbox;
                inbox.Open(FolderAccess.ReadWrite);
                for (int i = 0; i < inbox.Count; i++)
                {
                    var propertyMessage = inbox.Fetch(new[] { i }, MessageSummaryItems.Flags);
                    if (!propertyMessage[0].Flags.Value.HasFlag(MessageFlags.Seen))
                    {
                        GetMessage(i, inbox, email, messageHandler);
                        inbox.AddFlags(i, MessageFlags.Seen, true);
                    }
                }
            }
        }
        public void GetMessage(int indexMessage, IMailFolder inbox, EmailUtils emailUtils, IMessageHandler messageHandler)
        {
            MimeMessage emailMessage = inbox.GetMessage(indexMessage);

            this.from = emailMessage.From.OfType<MailboxAddress>().Single().Address;
            this.subject = emailMessage.Subject;


            GetAttachmentFile(emailMessage, messageHandler, indexMessage);

        }
        public void GetAttachmentFile(MimeMessage message, IMessageHandler messageHandler, int indexMessage)
        {
            foreach (var attachment in message.Attachments)
            {
                MemoryStream ms = new MemoryStream();
                using (MemoryStream stream = ms)
                {
                    this.fileName = attachment.ContentDisposition.FileName;

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
                    if (!CheckFormat(attachment))
                    {
                        Boolean createDisk = email.AddFileToDist(stream, fileName, patchToDisk);
                        if (createDisk)
                        {
                            string error = string.Format("Файл был пропущен , формат не подходит для обработки.");
                            email.sendMessage(from, error, uri, 587, login, password, patchToDisk + fileName);
                            logger.Error(String.Format("Ошибка : {0} Отправитель {1} , тема письма {2}", error, from, subject));
                            continue;
                        }
                        else
                        {
                            logger.Error(String.Format("Файл не был добавлен на диск название файла {0}.Отправитель {1} , тема письма {2} ", fileName, from, subject));
                            continue;
                        }
                    }
                    ExcelToJSON(stream, messageHandler, indexMessage);
                }
            }
        }
        public void ExcelToJSON(MemoryStream stream, IMessageHandler messageHandler, int indexMessage)
        {
            try
            {
                reader = ExcelReaderFactory.CreateReader(stream);
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        //UseHeaderRow = true
                        UseHeaderRow = false,
                        FilterRow = rowReader => rowReader.Depth >= startLine
                    }
                };
                List<rowSetting> rowSettings = GetSetting(indexMessage);
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
                    for (int j = 0; j < rowSettings.Count(); j++)
                    {
                        newRow[dt.Columns[j].ColumnName] = item[dataTable.Columns[rowSettings[j].numberCol - 1]];
                    }
                    dt.Rows.Add(newRow);
                }
                CreateESBMessage(JsonConvert.SerializeObject(dt), messageHandler, subject, from);
            }catch(Exception ex)
            {
                string error = string.Format("Ошибка преобразования файла " + ex.Message);
                //email.sendMessage(from + ";" + responsiblePerson, error, uri, 587, login, password);
                email.sendMessage(from, error, uri, 587, login, password);
                email.sendMessage(responsiblePerson, error, uri, 587, login, password);
                logger.Error(String.Format("Ошибка : {0} Отправитель {1} , тема письма {2}", error, from, subject));
            }
        }
        public List<rowSetting> GetSetting(int indexMessage)
        {
            FileStream openSettings = null;
            try
            {
                openSettings = File.Open(patchSetting + from + formatSetting, FileMode.Open, FileAccess.Read);
            }
            catch (Exception ex)
            {
                string error = String.Format(("Невозможно открыть настройки для отправителя {0} , тема письма {1}. \n Описание ошибка : {2} \n Trace : {3}"), from, subject, ex.Message, ex.StackTrace);
                email.sendMessage(responsiblePerson, error, uri, 587, login, password);
                throw new Exception(error);
            }
            StreamReader sr = new StreamReader(openSettings);

            JObject settingToProvider = JObject.Parse(sr.ReadToEnd());
            openSettings.Close();
            List<rowSetting> rowSettings = GetSettingsToRows(settingToProvider);

            if (rowSettings.Count < 0)
            {
                //logger.Warning(string.Format("Настройки не найдены для отправителя {0} , тема письма : {1}", from, subject));
                //inbox.AddFlags(i, MessageFlags.Seen, true);
                //  sr.Close();
                //continue;
                inbox.AddFlags(indexMessage, MessageFlags.Seen, true);
                string error = string.Format("Настройки не найдены для отправителя {0} , тема письма : {1}", from, subject);
                email.sendMessage(responsiblePerson, error, uri, 587, login, password);
                throw new Exception(error);
            }
            //sr.Close();

            startLine = (JsonUtils.IntValue(settingToProvider, "СхемаЗагрузки.НачальнаяСтрокаВФайле") - 1);
            sheetNumber = (JsonUtils.IntValue(settingToProvider, "СхемаЗагрузки.НомерЛистаВФайле") - 1);
            idObject = (JsonUtils.StringValue(settingToProvider, "СхемаЗагрузки.СсылкаНаСхему"));
            return rowSettings;
        }
        public bool CheckFormat(MimeEntity attachment)
        {
            var ext = Regex.Match(attachment.ContentType.Name, "[^.]+$").Value;
            switch (ext)
            {
                case "xlsx":
                    break;
                case "xls":
                    break;
                default:
                    return false;
            }
            return true;
        }
        public List<rowSetting> GetSettingsToRows(JObject jObject)
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
        public class rowSetting
        {
            public int numberCol { get; set; }
            public string viewCol { get; set; }
        }
        public void CreateESBMessage(string jsonBody, IMessageHandler messageHandler, string subject, string from)
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
            esbMessage.SetPropertyWithValue("idObject", idObject);

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
            email = new EmailUtils();
        }
    }
}
