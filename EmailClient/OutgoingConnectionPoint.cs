using ESB_ConnectionPoints.PluginsInterfaces;
using System.Threading;

namespace EmailClient
{
    class OutgoingConnectionPoint : IStandartOutgoingConnectionPoint
    {
        private readonly ILogger logger;
        private string uidGoogle;

        public OutgoingConnectionPoint(string jsonSettings, IServiceLocator serviceLocator)
        {
            logger = serviceLocator.GetLogger(GetType());
        }
        public void Cleanup()
        {
            throw new System.NotImplementedException();
        }

        public void Dispose()
        {
            throw new System.NotImplementedException();
        }

        public void Initialize()
        {
            throw new System.NotImplementedException();
        }

        public void Run(IMessageSource messageSource, IMessageReplyHandler replyHandler, CancellationToken ct)
        {
            throw new System.NotImplementedException();
        }
    }
}
