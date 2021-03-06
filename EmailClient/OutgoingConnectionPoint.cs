using ESB_ConnectionPoints.PluginsInterfaces;
using ESB_ConnectionPoints.Utils;
using ExcelDataReader;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;

namespace EmailClient
{
    class OutgoingConnectionPoint : IStandartOutgoingConnectionPoint
    {
        private readonly ILogger logger;
        private bool debugMode;
        private readonly string urlPatch;
        private int timeOut;
        private string pathCatalog;
        string replyClassId, formatSetting, idObject;
        int startToRow = 0, numberList = 0;
        List<rowSetting> rowSettings = null;



        public OutgoingConnectionPoint(string jsonSettings, IServiceLocator serviceLocator)
        {

            logger = new Logger(serviceLocator.GetLogger(GetType()), debugMode, "REST клиент");
            urlPatch = @"https://drive.google.com/uc?export=download&id=";
            ParseSettings(jsonSettings);
        }
        public async void Run(IMessageSource messageSource, IMessageReplyHandler replyHandler, CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                Message message = null;
                try
                {
                    message = messageSource.PeekLockMessage(ct, 10000);
                }
                catch (Exception ex)
                {
                    logger.Error(string.Format("Ошибка получения сообщения из очереди id : {0}. Описание ошибки : {1}", message.Id, ex.Message));
                }
                if (!(message == null))
                {
                    try
                    {
                        switch (message.Type)
	                        {
                                case "GDRV" : 
                                    byte[] body = Encoding.UTF8.GetBytes(ExcelToJSON(DowloadDocument(message, messageSource, logger)));
                                    Message replyMessage = new Message()
                                    {
                                        Body = body,
                                        ClassId = replyClassId,
                                        Id = Guid.NewGuid(),
                                        Type = "DTP"
                                    };
                                    replyMessage.SetPropertyWithValue("idObject", idObject);

                                    if (!replyHandler.HandleReplyMessage(replyMessage))
                                    {
                                        CompletePeeklock(logger, messageSource, message.Id, MessageHandlingError.UnknowError, "Ошибка отправки ответного сообщения");
                                    }
                                    CompletePeeklock(logger, messageSource, message.Id);
                                    break;
                                case "STG" :
                                    string[] arAddres = message.GetPropertyValue<string>("ArrayAddress").Split(';');
                                    string[] arSubject = message.GetPropertyValue<string>("ArraySubject").Split(';');
                                    foreach (string Address in arAddres)
                                    {
                                        await fileUtils.WriteSetting(message.Body, pathCatalog, (pathCatalog + Address + formatSetting),formatSetting);
                                    }
                                    foreach (string Subject  in arSubject)
                                    {
                                        await fileUtils.WriteSetting(message.Body, pathCatalog, (pathCatalog + Subject + formatSetting), formatSetting);
                                    }
                                    CompletePeeklock(logger, messageSource, message.Id);
                                    break;
                                //case "LTR" :

		                        default:
                                break;
	                        }
                       /* if (message.Type == "GDRV")
                        {
                            byte[] body = Encoding.UTF8.GetBytes(ExcelToJSON(DowloadDocument(message, messageSource, logger)));
                            Message replyMessage = new Message()
                            {
                                Body = body,
                                ClassId = replyClassId,
                                Id = Guid.NewGuid(),
                                Type = "DTP"
                            };
                            replyMessage.SetPropertyWithValue("idObject", idObject);

                            if (!replyHandler.HandleReplyMessage(replyMessage))
                            {
                                //logger.Error("Ошибка отправки ответного сообщения");
                                CompletePeeklock(logger, messageSource, message.Id, MessageHandlingError.UnknowError, "Ошибка отправки ответного сообщения");
                            }
                            CompletePeeklock(logger, messageSource, message.Id);
                        }*/

                        //else if (message.Type == "STG")
                        //{
                        //    string[] arAddres = message.GetPropertyValue<string>("ArrayAddress").Split(';');
                        //    string[] arSubject = message.GetPropertyValue<string>("ArraySubject").Split(';');
                        //    foreach (string Address in arAddres)
                        //    {
                        //        await fileUtils.WriteSetting(message.Body, pathCatalog, (pathCatalog + Address + formatSetting),formatSetting);
                        //    }
                        //    foreach (string Subject  in arSubject)
                        //    {
                        //        await fileUtils.WriteSetting(message.Body, pathCatalog, (pathCatalog + Subject + formatSetting), formatSetting);
                        //    }
                        //    CompletePeeklock(logger, messageSource, message.Id);
                        //}
                    }
                    catch (Exception ex)
                    {
                        CompletePeeklock(logger, messageSource, message.Id, MessageHandlingError.UnknowError, ex.Message);
                        continue;
                    }
                }
            }
        }
        //public void WriteSetting(string uid, byte[] body, string formatFile)
        //{
        //    var path = patchFile + uid + formatFile;
        //    DirectoryInfo dirInfo = new DirectoryInfo(patchFile);
        //    FileInfo fileInfo = new FileInfo(path);
        //    //Создаем каталог для файла
        //    if (!dirInfo.Exists)
        //        dirInfo.Create();
        //    //Удаляем файл перед записью
        //    if(fileInfo.Exists)
        //    {
        //        fileInfo.Delete();
        //        //using (var fs = new FileStream(patchFile + uid + formatFile, FileMode.Truncate))
        //        //{
        //        //}
        //    }
        //    using (FileStream fs = new FileStream(patchFile + uid + formatFile, FileMode.OpenOrCreate))
        //    {

        //        byte[] array = body;

        //        fs.Write(array, 0, array.Length);
        //    }
        //}
        public string ExcelToJSON(MemoryStream ms)
        {
            IExcelDataReader dataReader = ExcelReaderFactory.CreateReader(ms);
            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = false,
                    FilterRow = rowReader => rowReader.Depth > startToRow
                }
            };
            DataSet dataSet = dataReader.AsDataSet(conf);
            DataTable dataTable = dataSet.Tables[numberList];
            DataTable dt = new DataTable();

            foreach (var item in rowSettings)
            {
                DataColumn column = new DataColumn()
                {
                    ColumnName = item.viewCol,
                    DataType = dataTable.Columns[item.numberCol - 1].DataType,//System.Type.GetType("System.String"),
                    Expression = dataTable.Columns[item.numberCol - 1].Expression,
                    ColumnMapping = dataTable.Columns[item.numberCol - 1].ColumnMapping
                };
                var a = dataTable.Columns[item.numberCol - 1].ColumnName;
                dt.Columns.Add(column);
            }
            foreach (DataRow row in dataTable.Rows)
            {
                DataRow newRow = dt.NewRow();
                for (int i = 0; i < rowSettings.Count(); i++)
                {
                    newRow[dt.Columns[i].ColumnName] = row[dataTable.Columns[rowSettings[i].numberCol - 1]];
                }
                dt.Rows.Add(newRow);
            }
            return JsonConvert.SerializeObject(dt);
        }
        public MemoryStream DowloadDocument(Message message, IMessageSource messageSource, ILogger logger)
        {
            if (!GetSettings(message.GetPropertyValue<string>("Id")))
            {
                throw new Exception("Ошибка получения настроек");
            }
            using (HttpClient httpClient = new HttpClient())
            {
                var GetTask = httpClient.GetAsync(urlPatch + message.GetPropertyValue<string>("Id"));
                GetTask.Wait(timeOut);

                if (!GetTask.Result.IsSuccessStatusCode)
                {
                    CompletePeeklock(logger, messageSource, message.Id, MessageHandlingError.UnknowError, GetTask.Result.StatusCode.ToString());
                    throw new Exception("Ошибка отправки запроса серверу");
                }
                MemoryStream fs = new MemoryStream();
                var ResponseTask = GetTask.Result.Content.CopyToAsync(fs);
                ResponseTask.Wait(timeOut);
                return fs;
            }
        }
        public bool GetSettings(string uid)
        {
            JObject setting;
            try
            {
               setting = JObject.Parse(fileUtils.GetSetting(pathCatalog + uid));
            }
            catch (Exception ex)
            {
                logger.Error("Ошибка разбора JSON настройки поставщика " + ex.Message);
                return false;
            }
            this.startToRow = JsonUtils.IntValue(setting, "СхемаЗагрузки.НачальнаяСтрокаВФайле");
            this.numberList = JsonUtils.IntValue(setting, "СхемаЗагрузки.НомерЛистаВФайле");
            this.idObject = JsonUtils.StringValue(setting, "СхемаЗагрузки.СсылкаНаСхему");
            foreach (var item in setting["СхемаЗагрузки"]["ЗагружаеммыеПоля"])
            {
                rowSetting rowSetting = new rowSetting
                {
                    numberCol = (int)item["НомерКолонкиВФайле"],
                    viewCol = item["ВидКолонки"].ToString()
                };
                this.rowSettings.Add(rowSetting);
            }
            return true;
        }
        public void CompletePeeklock(ILogger logger, IMessageSource messageSource, Guid Id, MessageHandlingError messageHandlingError, string errorMessage)
        {
            logger.Error(string.Format("Ошибка отправки сообщения , подробности : {0}", errorMessage));
            messageSource.CompletePeekLock(Id, messageHandlingError, errorMessage);

        }
        private void CompletePeeklock(ILogger logger, IMessageSource messageSource, Guid Id)
        {
            messageSource.CompletePeekLock(Id);
            logger.Debug("Сообщение обработано");
        }
        public void ParseSettings(string jsonSettings)
        {
            if (string.IsNullOrEmpty(jsonSettings))
                throw new Exception("Не заданы параметры <jsonSettings>");

            JObject jObject;
            try
            {
                jObject = JObject.Parse(jsonSettings);
            }
            catch (Exception ex)
            {

                throw new Exception("Не удалось разобрать строку настроек JSON ! Ошибка : " + ex.Message);
            }
            debugMode = JsonUtils.BoolValue(jObject, "debugMode");
            timeOut = JsonUtils.IntValue(jObject, "timeOut", 10);
            pathCatalog = JsonUtils.StringValue(jObject, "pathCatalog", @"C:\ProgramData\adapterFile");
            replyClassId = JsonUtils.StringValue(jObject, "replyClassId");
            formatSetting = JsonUtils.StringValue(jObject, "formatSetting", ".JSON");
        }
        public void Initialize()
        {
        }
        public void Cleanup()
        {

        }
        public void Dispose()
        {

        }
    }
}

class rowSetting
{
    public int numberCol { get; set; }
    public string viewCol { get; set; }
}
