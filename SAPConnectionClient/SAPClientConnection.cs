using System;
using System.ServiceModel;

namespace SAPConnectionClient
{
    [ServiceBehavior(InstanceContextMode = InstanceContextMode.Single, ConcurrencyMode = ConcurrencyMode.Multiple, IncludeExceptionDetailInFaults = true)]
    public class SAPClientConnection : System.ServiceModel.ClientBase<ISapService> , ISapService

    {

    }
}
