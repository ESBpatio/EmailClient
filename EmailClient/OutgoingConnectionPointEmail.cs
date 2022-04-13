using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ESB_ConnectionPoints.PluginsInterfaces;
using ESB_ConnectionPoints.Utils;
using ExcelDataReader;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace EmailClient
{
    class OutgoingConnectionPointEmail : IStandartOutgoingConnectionPoint
    {
        private readonly ILogger logger;
        private bool debugMode;
        private string settingCatalog, formatSetting;
        private readonly string urlPatch;
        private int timeOut;
        //private string pathCatalog;
        string replyClassId, idObject, YNP , url;
        int startToRow = 0, numberList = 0;
        //List<rowSetting> rowSettings = new List<rowSetting>();

        public OutgoingConnectionPointEmail(string jsonSettings, IServiceLocator serviceLocator)
        {
            logger = new Logger(serviceLocator.GetLogger(GetType()), debugMode, "Rest client");
            urlPatch = @"https://drive.google.com/uc?export=download&id=";
            ParseSettings(jsonSettings);
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

        public async void Run(IMessageSource messageSource, IMessageReplyHandler replyHandler, CancellationToken ct)
        {
            while(!ct.IsCancellationRequested)
            {
                Message message = null;
                try
                {
                    message = messageSource.PeekLockMessage(ct, 10000);
                }
                catch (Exception ex)
                {

                    logger.Error(string.Format("Не удалось получить сообщения из очереди id сообщения : {0} " +
                        "\n Описание ошибки : {1}", message.Id, ex));
                }
                if(!(message == null))
                {
                    try
                    {
                        switch (message.Type)
                        {
                            case "STG":
                                string[] arAddress = null, arSubject = null;
                                List<string> listAddress = new List<string>();
                                if (message.HasProperty("ArrayAddress"))
                                {
                                    arAddress = message.GetPropertyValue<string>("ArrayAddress").Split(';');
                                    listAddress = GetAddressList(arAddress);
                                }
                                if (message.HasProperty("ArraySubject"))
                                {
                                    arSubject = message.GetPropertyValue<string>("ArraySubject").Split(';');
                                    for(int i = 0; i < arSubject.Count(); i++)
                                    {
                                        string patch = settingCatalog + arSubject[i];
                                        await fileUtils.WriteSetting(message.Body, patch, patch + @"\" + arAddress[i] + formatSetting, formatSetting);
                                        listAddress.Remove(arAddress[i]);
                                    }
                                }
                                foreach (string address in listAddress)
                                {
                                    await fileUtils.WriteSetting(message.Body, settingCatalog, (settingCatalog + address + formatSetting), formatSetting);
                                }
                                break;
                            case "GDRV":
                                {
                                    JObject outgoingBody;
                                    try
                                    {
                                        outgoingBody = JObject.Parse(Encoding.UTF8.GetString(message.Body));
                                    }
                                    catch (Exception)
                                    {

                                        CompletePeeklock(logger, messageSource, message.Id, MessageHandlingError.InvalidMessageFormat, "Не удалось распарсить тело сообщения");
                                        break;
                                    }
                                    List<rowSetting> rowSettings = GetSetting(outgoingBody);
                                    if (rowSettings.Count == 0)
                                    {
                                        CompletePeeklock(logger, messageSource, message.Id, MessageHandlingError.UnknowError, "Не удалось получить настройки из тела сообщения");
                                        break;
                                    }
                                    DirectoryInfo catalog =  fileUtils.CreateCatalog($"{settingCatalog}dowloadGoogleFile");
                                    logger.Debug($"Получен каталог на сервере {catalog.FullName}");
                                    var pathFile = fileUtils.DowloadFileFromURLToPath(GetGoogleTableDownloadLinkFromUrl(this.url), catalog + @"\" + idObject + ".xlsx");
                                    if (pathFile == null)
                                    {
                                        CompletePeeklock(logger, messageSource, message.Id, MessageHandlingError.UnknowError, "Не удалось сохранить файл в каталог");
                                        break;
                                    }                                      
                                    logger.Debug($"Создан файл на сервере для обработки {pathFile.FullName}");
                                    using (MemoryStream ms = fileUtils.GetFileStream(pathFile.FullName))
                                    {
                                        var c = ExcelToJSON(ms, rowSettings);
                                        Message replyMessage = new Message()
                                        {
                                            Body = Encoding.UTF8.GetBytes(c),
                                            ClassId = replyClassId,
                                            Id = Guid.NewGuid(),
                                            Type = "DTP"
                                        };
                                        replyMessage.SetPropertyWithValue("YNP", YNP);
                                        replyMessage.SetPropertyWithValue("API", false);
                                        replyMessage.SetPropertyWithValue("idObject", idObject);
                                        replyHandler.HandleReplyMessage(replyMessage);
                                    } 
                                    break;
                                }
                            case "SDRV":
                                {
                                    //settingCatalog = settingCatalog + @"\GoogleSettings\";
                                    await fileUtils.WriteSetting(message.Body, settingCatalog, (settingCatalog  + message.GetPropertyValue<string>("Id") + formatSetting), formatSetting);
                                    break;
                                }
                            default:                               
                                break;
                        }
                        CompletePeeklock(logger, messageSource, message.Id);
                    }
                    catch (Exception ex)
                    {

                        CompletePeeklock(logger, messageSource, message.Id, MessageHandlingError.UnknowError, ex.Message);
                        continue;
                    }
                }
            }
        }
        public List<string> GetAddressList(string[] arAddress)
        {
            List<string> listAddress = new List<string>();
            foreach (string item in arAddress)
            {
                listAddress.Add(item);
            }
            return listAddress;
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
            settingCatalog = JsonUtils.StringValue(jObject, "pathCatalog", @"C:\ProgramData\adapterFile");
            replyClassId = JsonUtils.StringValue(jObject, "replyClassId");
            formatSetting = JsonUtils.StringValue(jObject, "formatSetting", ".JSON");
        }
        public string ExcelToJSON(MemoryStream ms, List<rowSetting> rowSettings)
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
        //public bool GetSettings(string uid)
        //{
        //    //List<rowSetting> rowSettings = new List<rowSetting>();
        //    JObject setting;
        //    try
        //    {
        //        setting = JObject.Parse(fileUtils.GetSetting($"{settingCatalog}{uid}{formatSetting}")) ;
        //    }
        //    catch (Exception ex)
        //    {
        //        logger.Error("Ошибка разбора JSON настройки поставщика " + ex.Message);
        //        return false;
        //    }
        //    this.startToRow = JsonUtils.IntValue(setting, "СхемаЗагрузки.НачальнаяСтрокаВФайле");
        //    this.numberList = JsonUtils.IntValue(setting, "СхемаЗагрузки.НомерЛистаВФайле");
        //    this.idObject = JsonUtils.StringValue(setting, "СхемаЗагрузки.СсылкаНаСхему");
        //    foreach (var item in setting["СхемаЗагрузки"]["ЗагружаеммыеПоля"])
        //    {
        //        rowSetting rowSetting = new rowSetting
        //        {
        //            numberCol = (int)item["НомерКолонкиВФайле"],
        //            viewCol = item["ВидКолонки"].ToString()
        //        };
        //        this.rowSettings.Add(rowSetting);
        //    }
        //    return true;
        //}
        public List<rowSetting> GetSetting(JObject outgoingBody)
        {
            try
            {
                    JObject settingsToProvider = outgoingBody;
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
                    this.startToRow = JsonUtils.IntValue(settingsToProvider, "СхемаЗагрузки.НачальнаяСтрокаВФайле") - 1;
                    this.numberList = JsonUtils.IntValue(settingsToProvider, "СхемаЗагрузки.НомерЛистаВФайле") - 1;
                    this.idObject = JsonUtils.StringValue(settingsToProvider, "НастройкиПолей.СсылкаНаСхему");
                    this.YNP = JsonUtils.StringValue(settingsToProvider, "НастройкиПолей.УНПКонтрагента");
                    this.url = JsonUtils.StringValue(settingsToProvider, "НастройкиПолей.СсылкаДляПолученияПрайса");
                    return rowSettings;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        //public MemoryStream DowloadDocument(Message message, IMessageSource messageSource, ILogger logger)
        //{
        //    if (!GetSettings(message.GetPropertyValue<string>("Id")))
        //    {
        //        throw new Exception("Ошибка получения настроек");
        //    }
        //    using (HttpClient httpClient = new HttpClient())
        //    {
        //        //var GetTask = httpClient.GetAsync(urlPatch + message.GetPropertyValue<string>("Id"));
        //        var GetTask = httpClient.GetAsync($"https://docs.google.com/spreadsheets/d/1z3gsR6UCsQFgfiRunacp5eoxTHOb9MHYF3V2WJnl5_E/export?format=xlsx&id=1z3gsR6UCsQFgfiRunacp5eoxTHOb9MHYF3V2WJnl5_E");
        //        GetTask.Wait(timeOut);

        //        if (!GetTask.Result.IsSuccessStatusCode)
        //        {
        //            CompletePeeklock(logger, messageSource, message.Id, MessageHandlingError.UnknowError, GetTask.Result.StatusCode.ToString());
        //            throw new Exception("Ошибка отправки запроса серверу");
        //        }
        //        MemoryStream fs = new MemoryStream();
        //        var ResponseTask = GetTask.Result.Content.CopyToAsync(fs);
        //        ResponseTask.Wait(timeOut);
        //        return fs;
        //    }
        //}
        public static string GetGoogleTableDownloadLinkFromUrl(string url)
        {
            int index = url.IndexOf("/d/");
            int closingIndex;
            if (index > 0)
            {
                index += 3;
                closingIndex = url.IndexOf("/edit#", index);
                if (closingIndex < 0)
                    closingIndex = url.Length;
            }
            else
            {
                index = url.IndexOf("spreadsheets/d/");
                if (index < 0) // url is not in any of the supported forms
                    return string.Empty;

                index += 7;

                closingIndex = url.IndexOf('/', index);
                if (closingIndex < 0)
                {
                    closingIndex = url.IndexOf('?', index);
                    if (closingIndex < 0)
                        closingIndex = url.Length;
                }
            }
            return string.Format("https://docs.google.com/spreadsheets/d/{0}/export?format=xlsx&id={0}", url.Substring(index, closingIndex - index));
        }
    }
}
