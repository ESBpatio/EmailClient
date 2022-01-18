using ESB_ConnectionPoints.PluginsInterfaces;
using ESB_ConnectionPoints.Utils;
using System.Collections.Generic;

namespace EmailClient
{
    public sealed class IngoingConnectionPointFactory : IIngoingConnectionPointFactory
    {
        public IIngoingConnectionPoint Create(Dictionary<string, string> parameters, IServiceLocator serviceLocator)
        {
            //return (IIngoingConnectionPoint)new IngoingConnectionPoint(parameters.GetStringParameter("Настройки в формате JSON", true, ""), serviceLocator);
            return (IIngoingConnectionPoint)new IngointConnectionPointEmail(parameters.GetStringParameter("Настройки в формате JSON", true, ""), serviceLocator);

        }
    }
}
