using ESB_ConnectionPoints.PluginsInterfaces;
using System.Collections.Generic;
using ESB_ConnectionPoints.Utils;

namespace EmailClient
{
    public sealed class IngoingConnectionPointFactory : IIngoingConnectionPointFactory
    {
        public IIngoingConnectionPoint Create(Dictionary<string, string> parameters, IServiceLocator serviceLocator)
        {
            return (IIngoingConnectionPoint)new IngoingConnectionPoint(parameters.GetStringParameter("Настройки в формате JSON", true, ""), serviceLocator);
        }
    }
}
