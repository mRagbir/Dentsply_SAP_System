using System.Collections.Generic;
using System.Data;
using System.ServiceModel;

namespace Dentsply_SAP_Transactions_Service
{
    [ServiceContract(Namespace = "Dentsply_SAP_Transactions_Service")]
    public interface ISapService
    {

        [OperationContract]
        string Test(string value);

        // TODO: Add your service operations here

        [OperationContract]
        string PerformSAPHierarchyTransaction(string[] args);

        [OperationContract]
        string TestSAPConnection(string args);

        [OperationContract]
        bool ValidateMaterialNumber(string args);

        [OperationContract]
        List<string> ParseUDIBarcode(string args);

        [OperationContract]
        DataTable GetUDISpecialCases();

        [OperationContract]
        DataTable GetMaterialInfoUsingMaterialNumber(string args);


        #region Check if serials exist

        [OperationContract]
        DataTable CheckIfKitSerialExists(string materialNumber,string serial);

        [OperationContract]
        DataTable CheckIfSensorSerialExists(string materialNumber, string serial);

        [OperationContract]
        DataTable CheckIfCableSerialExists(string materialNumber, string serial);

        [OperationContract]
        DataTable CheckIfRemoteSerialExists(string materialNumber, string serial);

        #endregion

    }



}
