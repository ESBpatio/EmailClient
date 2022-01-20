using ESB_ConnectionPoints.PluginsInterfaces;
using ESB_ConnectionPoints.Utils;
using System.Collections.Generic;

namespace EmailClient
{
    class OutgoingConnectionPointFactory : IOutgoingConnectionPointFactory
    {
        public IOutgoingConnectionPoint Create(Dictionary<string, string> parameters, IServiceLocator serviceLocator)
        {
            //return (IOutgoingConnectionPoint)new OutgoingConnectionPoint(
            //    parameters.GetStringParameter("Настройки в формате JSON"), serviceLocator
            return (IOutgoingConnectionPoint)new OutgoingConnectionPointEmail(
    parameters.GetStringParameter("Настройки в формате JSON"), serviceLocator
            );
        }
    }
}
