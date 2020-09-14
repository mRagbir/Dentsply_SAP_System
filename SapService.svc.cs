
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text.RegularExpressions;
using WcfSAPService.Business_Logic;

namespace Dentsply_SAP_Transactions_Service
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select Service1.svc or Service1.svc.cs at the Solution Explorer and start debugging.
    [ServiceBehavior(InstanceContextMode = InstanceContextMode.Single, ConcurrencyMode = ConcurrencyMode.Multiple)]

    [KnownType(typeof(DataTable))]
    public class SapService : ISapService
    {
        public delegate void MonitorMessageDelegate(string aMessage);
        public event MonitorMessageDelegate Messages = null;

        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);


        public string Test(string value)
        {
            if (Messages != null)
                Messages(value);
            return string.Format($"From Dentsply_SAP_Transactions_Service -- You entered: {value}");
        }

        #region SAP Related

        /// <summary>
        /// SAP API needs 20 arguments to complete SAP hierarchy transaction
        /// (1) Order_Number as String
        /// (2) Material_Number As String
        /// (3) KitSerial_Global As String
        /// (4) Storage_Location As String, 'Storage location
        /// (5) TypeOfGR As String,'Determines the selection in the GR dropdown
        /// (6) bIsSensorNeeded As Boolean,'Determines if a sensor is needed in kit (Not needed if sensor # is Kit #)
        /// (7) bIsSpareCableNeeded As Boolean,'Determines if a spare cable is needed in kit (No spare for RMA kit)
        /// (8) bIsStarterKit As Boolean,'Determines if you are using a starter kit
        /// (9) bIsBasicKit As Boolean,'Determines if you are using a basic kit
        /// (10) bIsCableNeededOnSensor As Boolean,'Determines if a cable is needed on the sensor(eg: Starter kit will not use a cable as it is finished)
        /// (11) bIsChild1Checked As Boolean,'does checkbox 1 in hierarchy need to be checked ?
        /// (12) bIsChild2Checked As Boolean,'does checkbox 2 in hierarchy need to be checked ?
        /// (13) bIsChild3Checked As Boolean,'does checkbox 3 in hierarchy need to be checked ?
        /// (14) Sensor_MaterialNumber As String , 'Sensor Material #
        /// (15) Sensor_SerialNumber As String , 'Sensor serial #
        /// (16) Cable_MaterialNumber As String , 'Cable material #
        /// (17) CableSerial As String , 'Cable serial #
        /// (18) SpareCableSerial As String , 'Spare cable serial #
        /// (19) RemotePartNumber As String ,'Remote Part Number
        /// (20) Remote_SerialNumber As String ' Remote Serial Number
        /// </summary>
        /// <param name="args">Array of 20 arguments</param>
        /// <returns>bool true = all transactions performed , false = SAP error while performing transaction</returns>
        public string PerformSAPHierarchyTransaction(string[] args)
        {
            string returnResults = string.Empty;
            string[] tempArr;
            string[] finalArr;
            try
            {
                if (args.Length < 20)
                {
                    returnResults = "Incorrect number of args, please check if you have 20 args before using api";
                }
                else
                {
                    tempArr = new string[] { "PerformHierarchy" };
                    finalArr = tempArr.Concat(args).ToArray();


                    //get the exe location
                    string fileLocation = Environment.CurrentDirectory + @"\SAP_CONNECTION.exe";

                    System.Diagnostics.Process compiler = new System.Diagnostics.Process();
                    string AppPath = fileLocation;//ConfigurationManager.AppSettings["SAP_Connection"];
                    if (!File.Exists(AppPath)) throw new Exception("Cannot Find file location");
                    compiler = null;
                    string arguments = String.Join(" ", finalArr);
                    compiler = new System.Diagnostics.Process();
                    compiler.StartInfo.FileName = AppPath;
                    compiler.StartInfo.Verb = "runas";
                    compiler.StartInfo.Arguments = arguments;
                    compiler.StartInfo.UseShellExecute = false;
                    compiler.StartInfo.RedirectStandardOutput = true;
                    compiler.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
                    compiler.Start();
                    compiler.WaitForExit();
                    returnResults = compiler.StandardOutput.ReadToEnd();
                    //Console.WriteLine(results);
                    //Console.ReadLine();
                }
            }
            catch (Exception ex)
            {

                returnResults = ex.InnerException.ToString();

                logger.Error($"Error found : {ex.StackTrace}");
            }
            return returnResults;
        }

        /// <summary>
        /// Run a test to check if there is a valid SAp connection
        /// </summary>
        /// <param name="args">string parameter MUST BE -PerformConnectionTest-</param>
        /// <returns>string</returns>
        public string TestSAPConnection(string args)
        {
            string results = string.Empty;
            try
            {
                //get the exe location
                string fileLocation = Environment.CurrentDirectory + @"\SAP_CONNECTION.exe";

                System.Diagnostics.Process compiler = new System.Diagnostics.Process();
                string AppPath = fileLocation;//ConfigurationManager.AppSettings["SAP_Connection"];
                if (!File.Exists(AppPath)) throw new Exception("Cant find exe file location!");
                compiler = null;
                //string arguments = String.Join(",", args);
                compiler = new System.Diagnostics.Process();
                compiler.StartInfo.FileName = AppPath;
                compiler.StartInfo.Verb = "runas";
                compiler.StartInfo.Arguments = args;
                compiler.StartInfo.UseShellExecute = false;
                compiler.StartInfo.RedirectStandardOutput = true;
                compiler.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
                compiler.Start();
                compiler.WaitForExit();
                results = compiler.StandardOutput.ReadToEnd();
                //Console.WriteLine(results);
                //Console.ReadLine();

            }
            catch (Exception ex)
            {
                logger.Error($"Error found : {ex.StackTrace}");
                results = ex.Message;
            }

            return results;
        }

        public string TestSAPConnection()
        {
           return "Not Currently used";
        }

        #endregion

        #region Validate Material Methods

        /// <summary>
        /// Perform a check to the Production DB and determine if material exists
        /// </summary>
        /// <param name="materialNumber"></param>
        /// <returns>Boolean</returns>
        bool ISapService.ValidateMaterialNumber(string materialNumber)
        {
            bool bIsValid = false;

            try
            {
                //check if material number exists in the production utilities db
                DataTable dbTemp = DataToDB.SearchByMaterialNumber(materialNumber);
                dbTemp.TableName = "ValidateMaterialNumber";

                //check if DataTable came back with and info
                bIsValid = (dbTemp.Rows.Count > 0) ? true : false;

            }
            catch (Exception ex)
            {
                logger.Error($"Error : {ex.StackTrace}");
            }

            return bIsValid;
        }

        public DataTable GetMaterialInfoUsingMaterialNumber(string _materialNumber)
        {
            DataTable dtTemp = new DataTable();

            try
            {
                dtTemp = DataToDB.GetMaterialInfoUsingMaterialNumber(_materialNumber);
                dtTemp.TableName = "GetMaterialInfoUsingMaterialNumber";
            }
            catch (Exception e)
            {
                logger.Error($"Error : {e.StackTrace}");
            }

            return dtTemp;
        }


        #endregion

        #region Barcode parsing

        public List<string> ParseUDIBarcode(string _barcode)
        {
            List<string> lResults = new List<string>();
            List<string> lSpecialCases = new List<string>();

            try
            {
                //get the list of special cases
                DataTable dtTemp = GetUDISpecialCases();

                if (dtTemp.Rows.Count > 0)
                {
                    foreach (DataRow dr in dtTemp.Rows)
                    {
                        lSpecialCases.Add(dr["Material_Number"].ToString());
                    }
                }

                //step 1 remove all special characters
                string removeSpecialChars = Regex.Replace(_barcode, @"[+%$]", "");

                //step 2 remove the first 4 characters of string
                string removeFirst4 = removeSpecialChars.Remove(0, 4);

                //step 3 split string by the '/'
                string[] tempSpl = removeFirst4.Split('/');

                //step 4 remove last char from tempSpl[0]
                tempSpl[0] = tempSpl[0].Substring(0, tempSpl[0].Length - 1);

                //step 5 determine if the serial has an extra identifier according to the lSpecialCases
                if (lSpecialCases.Any(s => s.Contains(tempSpl[0])))
                {
                    tempSpl[1] = tempSpl[1].Substring(1);
                }


                lResults = tempSpl.ToList();

            }
            catch (Exception ex)
            {
                logger.Error($"Error : {ex.StackTrace}");
            }

            return lResults;
        }

        public DataTable GetUDISpecialCases()
        {
            DataTable dtTemp = new DataTable();
            try
            {
                 dtTemp = DataToDB.GetUDIBarcodeSpecialCases();
                dtTemp.TableName = "GetUDISpecialCases";
            }
            catch (Exception ex)
            {
                logger.Error($"Error : {ex.StackTrace}");
            }

            return dtTemp;
        }



        #endregion

        #region Check if serials exist

        public DataTable CheckIfKitSerialExists(string materialNumber, string serial)
        {
            DataTable dtTemp = new DataTable();

            try
            {
                dtTemp = DataToDB.dtCheckKitSerial(materialNumber,serial);
                dtTemp.TableName = "CheckIfKitSerialExists";
            }
            catch (Exception e)
            {
                logger.Error($"Error : {e.StackTrace}");
            }

            return dtTemp;
        }

        public DataTable CheckIfSensorSerialExists(string materialNumber, string serial)
        {
            DataTable dtTemp = new DataTable();

            try
            {
                dtTemp = DataToDB.dtCheckSensorSerial(materialNumber, serial);
                dtTemp.TableName = "CheckIfSensorSerialExists";
            }
            catch (Exception e)
            {
                logger.Error($"Error : {e.StackTrace}");
            }

            return dtTemp;
        }

        public DataTable CheckIfCableSerialExists(string materialNumber, string serial)
        {
            DataTable dtTemp = new DataTable();

            try
            {
                dtTemp = DataToDB.dtCheckCableSerial(materialNumber, serial);
                dtTemp.TableName = "CheckIfCableSerialExists";
            }
            catch (Exception e)
            {
                logger.Error($"Error : {e.StackTrace}");
            }

            return dtTemp;
        }

        public DataTable CheckIfRemoteSerialExists(string materialNumber, string serial)
        {
            DataTable dtTemp = new DataTable();

            try
            {
                dtTemp = DataToDB.dtCheckRemoteSerial(materialNumber, serial);
                dtTemp.TableName = "CheckIfRemoteSerialExists";
            }
            catch (Exception e)
            {
                logger.Error($"Error : {e.StackTrace}");
            }

            return dtTemp;
        }


        #endregion


    }
}
