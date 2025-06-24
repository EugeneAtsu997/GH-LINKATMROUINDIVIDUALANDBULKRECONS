using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace GH_LINK_ATM_ROU_INDIVIDUAL_AND_BULK_RECONS
{
    public class Program
    {
        static void Main(string[] args)
        {
            List<GhLinkReport> ghLinkReports = new List<GhLinkReport>();
            List<GhLinkInd> ghLinkInd = new List<GhLinkInd>();
            List<OpenItems> openItems = new List<OpenItems>();
            List<GhLinkReport> newGhLinkReports = new List<GhLinkReport>();


            try
            {
                Console.WriteLine("----------------Start--------------------");

                Console.WriteLine("-----------------------------------------");

                Console.WriteLine("Start Time ------------->   " + DateTime.Now);


                string inputInd = ConfigurationManager.AppSettings["inputInd"];
                string inputBulk = ConfigurationManager.AppSettings["inputBulk"];
                string generatedJson = ConfigurationManager.AppSettings["generatedJson"];
                string pivotInput = ConfigurationManager.AppSettings["pivotInput"];
                string outputTemplate = ConfigurationManager.AppSettings["outputTemplate"];
                string openItemsPath = ConfigurationManager.AppSettings["openItems"];
                string source = ConfigurationManager.AppSettings["source"];



                //Kill Excel Instances
                if (!UserFunctions.KillAllExcelInstances())
                {
                    Console.WriteLine("Unable to kill all excel instance");

                    Task.Factory.StartNew(() => UserFunctions.WriteLog(" ", " ", "Unable to kill all excel instance", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
                }

                Console.WriteLine(" ");

                Console.WriteLine("Excel instances killed successfully");
                Task.Factory.StartNew(() => UserFunctions.WriteLog(" ", "", "Excel instances killed successfully", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

                string message = string.Empty;
                string fileName = string.Empty;


                if (!UserFunctions.ReadAllFiles(inputInd, out List<string> filePath, "xlsx"))
                {
                    Console.WriteLine("No data found in source Path");

                    Task.Factory.StartNew(() => UserFunctions.WriteLog(" ", "", "No data found in source Path", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
                }

                Console.WriteLine(" ");

                Console.WriteLine("Data found in specified location");

                Task.Factory.StartNew(() => UserFunctions.WriteLog(" ", "", "Data found in specified location", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

                foreach (var item in filePath)
                {
                    fileName = Path.GetFileNameWithoutExtension(item);
                    string jsonInput = UserFunctions.ReadExcelToJson(item, generatedJson, fileName);



                    if (string.IsNullOrEmpty(jsonInput))
                    {
                        Console.WriteLine("Unable to read data from " + item);

                        Task.Factory.StartNew(() => UserFunctions.WriteLog(" ", fileName, "Data read from " + item + " unsuccessfully", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

                        //UserFunctions.MoveFile(item, workbookDone + Path.GetFileName(item));

                        Thread.Sleep(1500);

                        continue;
                    }

                    Console.WriteLine(" ");
                    Console.WriteLine("Data read from " + item + " successfully");

                    Task.Factory.StartNew(() => UserFunctions.WriteLog(" ", fileName, "Data read from " + item + " successfully", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

                    Console.WriteLine(fileName);

                    UserFunctions.ReadJson(jsonInput, out ghLinkReports);
                    Console.WriteLine("");


                }

                // Get Open Items
                UserFunctions.GetOpenItemsFile(openItemsPath, generatedJson, source, out openItems, out message);
                Console.WriteLine(message);

                // Merge Open items and new data
                UserFunctions.MergeOpenItemsNReport(ghLinkReports, openItems, out newGhLinkReports, out message);
                Console.WriteLine(message);

                UserFunctions.GetReconKey(newGhLinkReports, out ghLinkInd, out message);
                Console.WriteLine(message);

                UserFunctions.WriteToExcelSheet(new List<string> { JsonConvert.SerializeObject(ghLinkInd) }, outputTemplate, pivotInput, StaticVariables.SHEET1, StaticVariables.PIVOTDATA, out message);
                Console.WriteLine(message);
            }
            catch (Exception ex)
            {

                Task.Factory.StartNew(() => UserFunctions.WriteLog(" ", " ", ex.Message + "  || " + ex.StackTrace, "Error", string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
                Console.WriteLine(" ");
                Console.WriteLine("Exception -------------------->    " + ex.Message + "  || " + ex.StackTrace);
            }

            Console.WriteLine("");
            Console.WriteLine("Process Completed @ " + DateTime.Now);

            Thread.Sleep(150000);
        }
    }
}
