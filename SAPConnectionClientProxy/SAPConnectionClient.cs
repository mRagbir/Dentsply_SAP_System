using Dentsply_SAP_Transactions_Service;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace SAPConnectionClientProxy
{
    //[ServiceBehavior(InstanceContextMode = InstanceContextMode.Single, ConcurrencyMode = ConcurrencyMode.Multiple, IncludeExceptionDetailInFaults = true)]
    //[KnownType(typeof(DataTable))]
    public class SAPConnectionClient : System.ServiceModel.ClientBase<ISapService>, ISapService
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);


        public SAPConnectionClient(string endpointConfigurationName) : base(endpointConfigurationName)
        {

        }

        #region Tests

        public string Test(string value)
        {
            try
            {
                return Channel.Test(value);
            }
            catch (Exception e)
            {
                logger.Error($"Error : {e.StackTrace}");
                return e.Message;
            }
        }


        #endregion

        #region Check if serial exists


        /// <summary>
        /// Query DB to find out if serial was already used
        /// </summary>
        /// <param name="materialNumber">Cable Material Number</param>
        /// <param name="serial">Cable Serial Number</param>
        /// <returns></returns>
        public DataTable CheckIfCableSerialExists(string materialNumber, string serial)
        {
            try
            {
                return Channel.CheckIfCableSerialExists(materialNumber, serial);
            }
            catch (Exception e)
            {
                logger.Error($"Error : {e.StackTrace}");
                throw e;
            }
        }

        /// <summary>
        /// Query DB to find out if serial was already used
        /// </summary>
        /// <param name="materialNumber">Kit Material Number</param>
        /// <param name="serial">Kit Serial Number</param>
        /// <returns></returns>
        public DataTable CheckIfKitSerialExists(string materialNumber, string serial)
        {
            try
            {
                return Channel.CheckIfKitSerialExists(materialNumber, serial);
            }
            catch (Exception e)
            {
                logger.Error($"Error : {e.StackTrace}");
                throw e;
            }
        }

        /// <summary>
        /// Query DB to find out if serial was already used
        /// </summary>
        /// <param name="materialNumber">Remote Material Number</param>
        /// <param name="serial">Remote Serial Number</param>
        /// <returns></returns>
        public DataTable CheckIfRemoteSerialExists(string materialNumber, string serial)
        {
            try
            {
                return Channel.CheckIfRemoteSerialExists(materialNumber, serial);
            }
            catch (Exception e)
            {
                logger.Error($"Error : {e.StackTrace}");
                throw e;
            }
        }

        /// <summary>
        /// Query DB to find out if serial was already used
        /// </summary>
        /// <param name="materialNumber">Sensor Material Number</param>
        /// <param name="serial">Sensor Serial Number</param>
        /// <returns></returns>
        public DataTable CheckIfSensorSerialExists(string materialNumber, string serial)
        {
            try
            {
                return Channel.CheckIfSensorSerialExists(materialNumber, serial);
            }
            catch (Exception e)
            {
                logger.Error($"Error : {e.StackTrace}");
                throw e;
            }
        }

        #endregion



        /// <summary>
        /// Get all material info from the Production Utilities DB
        /// </summary>
        /// <param name="_materialNumber">Material number</param>
        /// <returns>Data Table</returns>
        public DataTable GetMaterialInfoUsingMaterialNumber(string _materialNumber)
        {
            try
            {
                return Channel.GetMaterialInfoUsingMaterialNumber(_materialNumber);
            }
            catch (Exception e)
            {
                logger.Error($"Error : {e.StackTrace}");
                throw  e;
            }
        }

        /// <summary>
        /// Get all material numbers that need special barcode parsing.
        /// This method is automatically used in the ParseUDIBarcode method
        /// </summary>
        /// <returns>Data Table</returns>
        public DataTable GetUDISpecialCases()
        {
            try
            {
                return Channel.GetUDISpecialCases();
            }
            catch (Exception e)
            {
                logger.Error($"Error : {e.StackTrace}");
                throw e;
            }
        }

        /// <summary>
        /// Parse a UDI barcode
        /// </summary>
        /// <param name="barcodeString">Barcode string</param>
        /// <returns>List<string> 0= Part Number , 1 = Serial Number , 2 = Date</string></returns>
        public List<string> ParseUDIBarcode(string barcodeString)
        {
            try
            {
                return Channel.ParseUDIBarcode(barcodeString);
            }
            catch (Exception e)
            {
                logger.Error($"Error : {e.StackTrace}");
                throw e;
            }

        }

        /// <summary>
        /// SAP API needs 20 arguments to complete SAP hierarchy transaction
        /// (1) Order_Number as String
        /// (2) Material_Number As String
        /// (3) KitSerial_Global As String
        /// (4) Storage_Location As String,
        /// (5) TypeOfGR As String,'Determines the selection in the GR dropdown
        /// (6) bIsSensorNeeded As Boolean,'Determines if a sensor is needed in kit (Not needed if sensor # is Kit #)
        /// (7) bIsSpareCableNeeded As Boolean,'Determines if a spare cable is needed in kit (No spare for RMA kit)
        /// (8) bIsStarterKit As Boolean,'Determines if you are using a starter kit
        /// (9) bIsBasicKit As Boolean,'Determines if you are using a basic kit
        /// (10) bIsCableNeededOnSensor As Boolean,'Determines if a cable is needed on the sensor(eg: Starter kit will not use a cable as it is finished)
        /// (11) bIsChild1Checked As Boolean,'does checkbox 1 in hierarchy need to be checked ?
        /// (12) bIsChild2Checked As Boolean,'does checkbox 2 in hierarchy need to be checked ?
        /// (13) bIsChild3Checked As Boolean,'does checkbox 3 in hierarchy need to be checked ?
        /// (14) Sensor_MaterialNumber As String ,
        /// (15) Sensor_SerialNumber As String ,
        /// (16) Cable_MaterialNumber As String ,
        /// (17) CableSerial As String ,
        /// (18) SpareCableSerial As String ,
        /// (19) RemotePartNumber As String ,
        /// (20) Remote_SerialNumber As String
        /// </summary>
        /// <param name="args">Array of 20 arguments</param>
        /// <returns>bool true = all transactions performed , false = SAP error while performing transaction</returns>
        public string PerformSAPHierarchyTransaction(string[] args)
        {
            try
            {
                return Channel.PerformSAPHierarchyTransaction(args);
            }
            catch (Exception e)
            {
                logger.Error($"Error : {e.StackTrace}");
                return e.Message;
            }
        }

        /// <summary>
        /// Run a test to check if there is a valid SAP connection
        /// </summary>
        /// <param name="args">string parameter MUST BE ~~> (PerformConnectionTest) </param>
        /// <returns>string</returns>
        public string TestSAPConnection(string args)
        {
            string results = string.Empty;
            try
            {
                results= Channel.TestSAPConnection(args);
                return results;
            }
            catch (Exception e)
            {
                logger.Error($"Error : {e.StackTrace}");
                return e.Message;
            }
        }

        /// <summary>
        /// Perform a check to the Production DB and determine if material exists
        /// </summary>
        /// <param name="args"></param>
        /// <returns>Boolean</returns>
        public bool ValidateMaterialNumber(string args)
        {
            bool bIsValid = false;

            try
            {
                bIsValid = Channel.ValidateMaterialNumber(args);
            }
            catch (Exception ex)
            {
                logger.Error($"Error : {ex.StackTrace}");
                throw new FaultException( ex.Message);
            }

            return bIsValid;



        }
    }
}
