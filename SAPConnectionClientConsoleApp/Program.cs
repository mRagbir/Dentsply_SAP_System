using SAPConnectionClientProxy;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SAPConnectionClientConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {

            try
            {
                SAPConnectionClient client = new SAPConnectionClient("SAPTCP");

                string[] argsForHierarchy = new string[] {"Order_Number","Material_Number","KitSerial","Storage_Location","TypeOfGR","bIsSensorNeeded","bIsSpareCableNeeded","bIsStarterKit","bIsBasicKit","bIsCableNeededOnSensor",
                                                          "bIsChild1Checked","bIsChild2Checked","bIsChild3Checked","Sensor_MaterialNumber","Sensor_SerialNumber","Cable_MaterialNumber","CableSerial","SpareCableSerial","RemotePartNumber","Remote_SerialNumber"};

                //connection test
                string results = string.Empty;

                ////determine is connection test passed
                //bool bIsPassed = Convert.ToBoolean( results.Split(':')[1].ToString());

                //if (bIsPassed)
                //{
                //    //try another transaction
                //    // string r = client.PerformSAPHierarchyTransaction(argsForHierarchy);
                //}

                string barcode = @"+D764B11180521/$+JL10001234/16D20200123+";
                //var x = client.


                DataTable dtTemp = client.GetMaterialInfoUsingMaterialNumber("100007060");


                Console.WriteLine(results);
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
