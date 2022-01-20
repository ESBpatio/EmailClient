using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ESB_ConnectionPoints.PluginsInterfaces;
using ESB_ConnectionPoints.Utils;
using Newtonsoft.Json.Linq;

namespace EmailClient
{
    class OutgoingConnectionPointEmail : IStandartOutgoingConnectionPoint
    {
        private readonly ILogger logger;
        private bool debugMode;
        private string settingCatalog, formatSetting;

        public OutgoingConnectionPointEmail(string jsonSettings, IServiceLocator serviceLocator)
        {
            logger = new Logger(serviceLocator.GetLogger(GetType()), debugMode, "Rest client");
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
            //timeOut = JsonUtils.IntValue(jObject, "timeOut", 10);
            settingCatalog = JsonUtils.StringValue(jObject, "pathCatalog", @"C:\ProgramData\adapterFile");
            //replyClassId = JsonUtils.StringValue(jObject, "replyClassId");
            formatSetting = JsonUtils.StringValue(jObject, "formatSetting", ".JSON");
        }
    }
}
