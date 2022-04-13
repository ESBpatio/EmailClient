using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using ESB_ConnectionPoints.PluginsInterfaces;
using ExcelDataReader;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace EmailClient
{
    class IngointConnectionPointEmail : IStandartIngoingConnectionPoint
    {
        private readonly ILogger logger;
        private readonly IMessageFactory messageFactory;
        private string Uri, Login, Password, ClassId, Type, FormatSetting, PatchSetting,
            From, Subject, PatchToDisk, ResponsiblePerson, FileName, ObjectId, YNP;
        private bool Ssl;
        private int Port, Timeout, StartLine, SheetNumber;
        //public IExcelDataReader dataReader;
        public EmailUtils email;
        //public IMailFolder inbox;

        public IngointConnectionPointEmail(string jsonSettings, IServiceLocator serviceLocator)
        {
            if(serviceLocator != null)
            {
                this.logger = serviceLocator.GetLogger(this.GetType());
                this.messageFactory = serviceLocator.GetMessageFactory();
            }
            email = new EmailUtils();
            this.ParseSettings(jsonSettings);
        }

        public void GetMessages(IMessageHandler messageHandler)
        {
            using (ImapClient client = new ImapClient())
            {
                client.Connect(Uri, Port, Ssl);
                client.Authenticate(Login, Password);
                if(client.IsAuthenticated)
                {
                    client.Inbox.Open(FolderAccess.ReadWrite);
                    IList<UniqueId> mailIds = client.Inbox.Search(SearchQuery.NotSeen);
                    foreach (var id in mailIds)
                    {
                        MimeMessage message = client.Inbox.GetMessage(id);
                        this.From = message.From.OfType<MailboxAddress>().Single().Address;
                        this.Subject = message.Subject;
                        string patch = PatchSetting + From + FormatSetting;
                        if (CheckSettingToSubject())
                            patch = PatchSetting + Subject + @"\" + From + FormatSetting;
                        List<rowSetting> rowSettings = GetSetting(patch);
                        foreach (var attachment in message.Attachments)
                        {
                            this.FileName = attachment.ContentDisposition.FileName;
                            if (!CheckFormat(attachment))
                            {
                                logger.Error("Файл был пропущен , неверный формат файла");
                                continue;
                            }
                            using (MemoryStream stream = new MemoryStream())
                            {                  
                                if(attachment is MessagePart)
                                {
                                    MessagePart message1 = (MessagePart)attachment;
                                    message1.Message.WriteTo(stream);
                                }
                                else
                                {
                                    MimePart message1 = (MimePart)attachment;
                                    message1.Content.DecodeTo(stream);
                                }
                                DataTable body = ExcelToJson(stream, rowSettings);
                                if (body == null)
                                {
                                    string error = string.Format("<p>Ошибка преобразования файла. \n  Имя файла {0} \n Отправитель {1}</p>", FileName, From);
                                    this.email.sendMessage(From, Uri, 587, Login, Password, "info.price@patio-minsk.by", "ESB Info", error, "Ошибка при загрузке вложения");
                                    this.email.sendMessage(ResponsiblePerson, Uri, 587, Login, Password, "info.price@patio-minsk.by", "ESB Info", error, "Ошибка при загрузке вложения");
                                    logger.Error(string.Format("Ошибка : {0} отправитель {1} , тема письма {2}", error, From, Subject));
                                    continue;
                                }
                                else
                                {
                                    messageHandler.HandleMessage(CreateESBMessage(JsonConvert.SerializeObject(body)));
                                }
                            }
                        }
                        client.Inbox.AddFlags(id, MessageFlags.Seen, true);
                    }                    
                }
            }
        }
      
        public Message CreateESBMessage(string body)
        {
            Message message = new Message
            {
                Body = Encoding.UTF8.GetBytes(body),
                ClassId = this.ClassId,
                Id = Guid.NewGuid(),
                Type = this.Type
            };
            message.SetPropertyWithValue("From", this.From);
            message.SetPropertyWithValue("Subject", this.Subject);
            message.SetPropertyWithValue("ObjectId", this.ObjectId);
            message.SetPropertyWithValue("YNP", this.YNP);
            return message;
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
            this.Uri = JsonUtils.StringValue(jObject, "ConfigurationServer.serverUri", "");
            this.Login = JsonUtils.StringValue(jObject, "ConfigurationServer.Login", "");
            this.Password = JsonUtils.StringValue(jObject, "ConfigurationServer.Password", "");
            this.Ssl = JsonUtils.BoolValue(jObject, "ConfigurationServer.Ssl", false);
            this.Port = JsonUtils.IntValue(jObject, "ConfigurationServer.Port", 993);
            this.Type = JsonUtils.StringValue(jObject, "MessageSettings.Type", "DTP");
            this.ClassId = JsonUtils.StringValue(jObject, "MessageSettings.ClassId", "0");
            this.Timeout = JsonUtils.IntValue(jObject, "TimeOut", 10);
            this.FormatSetting = JsonUtils.StringValue(jObject, "formatSetting", ".JSON");
            this.PatchSetting = JsonUtils.StringValue(jObject, "patchSetting", @"C:\Settings\");
            this.PatchToDisk = JsonUtils.StringValue(jObject, "patchToDisk", @"C:\tmp\");
            this.ResponsiblePerson = JsonUtils.StringValue(jObject, "responsiblePerson");
        }
     
        private bool CheckFormat(MimeEntity attachment)
        {
            switch (Regex.Match(attachment.ContentDisposition?.FileName, "[^.]+$").Value.ToLower())
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

        private DataTable ExcelToJson(MemoryStream stream, List<rowSetting> rowSettings)
        {
            try
            {
                IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
                ExcelDataSetConfiguration configurateReader = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = false,
                        FilterRow = rowReader => rowReader.Depth >= StartLine
                    }
                };
                DataSet dataSet = reader.AsDataSet(configurateReader);
                DataTable dataTable = dataSet.Tables[SheetNumber];
                DataTable dataTable1 = new DataTable();
                foreach (var item in rowSettings)
                {
                    DataColumn dataColumn = dataTable.Columns[item.numberCol - 1];
                    DataColumn dataColumn1 = new DataColumn()
                    {
                        ColumnName = item.viewCol,
                        DataType = dataColumn.DataType,
                        Expression = dataColumn.Expression,
                        ColumnMapping = dataColumn.ColumnMapping
                    };
                    dataTable1.Columns.Add(dataColumn1);
                }
                foreach (DataRow item in dataTable.Rows)
                {
                    DataRow row = dataTable1.NewRow();
                    for (int j = 0; j < rowSettings.Count(); j++)
                    {
                        row[dataTable1.Columns[j].ColumnName] = item[dataTable.Columns[rowSettings[j].numberCol - 1]];
                    }
                    dataTable1.Rows.Add(row);
                }
                return dataTable1;

            }
            catch (Exception ex)
            {
                return null;
            }
        }
        public List<rowSetting> GetSetting(string patch)
        {
            try
            {
                using (FileStream stream = File.Open(patch, FileMode.Open, FileAccess.Read))
                {
                    StreamReader sr = new StreamReader(stream);
                    JObject settingsToProvider = JObject.Parse(sr.ReadToEnd());
                    List<rowSetting> rowSettings = new List<rowSetting>();
                    foreach (var item in settingsToProvider["СхемаЗагрузки"]["ЗагружаеммыеПоля"])
                    {
                        rowSetting rowSetting = new rowSetting
                        {
                            numberCol = (int)item["НомерКолонкиВФайле"],
                            viewCol = item["ВидКолонки"].ToString()
                        };
                        rowSettings.Add(rowSetting);
                    }
                    this.StartLine = (JsonUtils.IntValue(settingsToProvider, "СхемаЗагрузки.НачальнаяСтрокаВФайле") - 1);
                    this.SheetNumber = (JsonUtils.IntValue(settingsToProvider, "СхемаЗагрузки.НомерЛистаВФайле") - 1);
                    this.ObjectId = (JsonUtils.StringValue(settingsToProvider, "НастройкиПолей.СсылкаНаСхему"));
                    this.YNP = (JsonUtils.StringValue(settingsToProvider, "НастройкиПолей.УНПКонтрагента"));
                    return rowSettings;
                }
            }
            catch(Exception ex)
            {
                return null;
            }
        }

        public bool CheckSettingToSubject()
        {
            if (File.Exists(PatchSetting + Subject + @"\"+ From + FormatSetting))
                return true;
            return false;
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
                    string error = string.Format("<h2> {0} \n Ошибка загрузки письма {1} \n Trace : {2} </h2>", "test", ex.Message, ex.StackTrace);
                    logger.Error(error);
                    //email.sendMessage(responsiblePerson, error, uri, 587, login, password);
                    email.sendMessage(ResponsiblePerson, Uri, 587, Login, Password, "info.price@patio-minsk.by", "ESB info", error, "Ошибка при загрузке вложения");
                }
                ct.WaitHandle.WaitOne(Timeout);
            }
        }
    }
    public class rowSetting
    {
        public int numberCol { get; set; }
        public string viewCol { get; set; }
    }
}
