using Aspose.Cells.Utility;
using Aspose.Cells;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices.ComTypes;

namespace GH_LINK_ATM_ROU_INDIVIDUAL_AND_BULK_RECONS
{
    public class UserFunctions
    {
        public static void WriteLog(string sescureId, string request, string response, string serviceame, string mfunctionName, [CallerMemberName] string callerName = "")
        {
            // Set mfunctionName to the callerName, which is an optional parameter
            mfunctionName = callerName;

            //Define the path to the log file, including the current date
            string logFilePath = "C:\\Logs\\" + serviceame + "\\";
            logFilePath = logFilePath + "Log-" + DateTime.Today.ToString("MM-dd-yyyy") + "." + "txt";

            try
            {
                //Create a FileStream for writing to the log file in append mode
                using (FileStream fileStream = new FileStream(logFilePath, FileMode.Append))
                {
                    FileInfo logFileInfo;

                    //Get information about the log file
                    logFileInfo = new FileInfo(logFilePath);

                    //Get information about the direction containing the log file
                    DirectoryInfo logDirInfo = new DirectoryInfo(logFileInfo.DirectoryName);

                    //If the directory does not exist, create it
                    if (!logDirInfo.Exists) logDirInfo.Create();

                    //Get access control information for the directory
                    DirectorySecurity dSecurity = logDirInfo.GetAccessControl();

                    //Add a permission rule to grant full control to the WorldSid(everyone)
                    dSecurity.AddAccessRule(new FileSystemAccessRule(new SecurityIdentifier(WellKnownSidType.WorldSid, null), FileSystemRights.FullControl, InheritanceFlags.ObjectInherit | InheritanceFlags.ContainerInherit, PropagationFlags.NoPropagateInherit, AccessControlType.Allow));

                    //Set the updated access control information for the directory
                    logDirInfo.SetAccessControl(dSecurity);

                    //Create a StreamWriter to write to the log file
                    StreamWriter log = new StreamWriter(fileStream);

                    // Create a StreamWriter to write to the log file
                    if (!logFileInfo.Exists)
                    {
                        _ = logFileInfo.Create();
                    }
                    else
                    {
                        // Write log entries to the log file
                        log.WriteLine(sescureId);
                        log.WriteLine(DateTime.UtcNow.ToString());
                        log.WriteLine(request);
                        log.WriteLine(response);
                        log.WriteLine(mfunctionName);
                        log.WriteLine("_____________________________________________________________________________________");

                        // Close the StreamWriter to flush and save the log entries
                        log.Close();
                    }
                    fileStream.Close();
                }
            }
            catch (Exception)
            {

            }
        }

        public static bool KillAllExcelInstances()
        {
            bool worked = false;
            try
            {
                Process[] process = Process.GetProcessesByName("Excel");

                foreach (Process p in process)
                {
                    if (!string.IsNullOrEmpty(p.ProcessName))
                    {
                        try
                        {
                            p.Kill();
                            worked = true;
                        }
                        catch
                        {

                        }
                    }
                }
                worked = true;
            }
            catch (Exception ex)
            {
                //Task.Factory.StartNew(() => WriteLog(" ", "", ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

            }
            return worked;
        }

        public static void ExcelUpdateAction(string workbookPath)
        {
            try
            {
                // Create a new instance of Microsoft Excel application
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application
                {
                    DisplayAlerts = false
                };

                // Open the specified Excel workbook
                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                // Get a reference to the worksheets in the workbook
                Microsoft.Office.Interop.Excel.Sheets worksheets = excelWorkbook.Worksheets;

                // Delete the worksheet named "Evaluation Warning"
                excelWorkbook.Sheets["Evaluation Warning"].Delete();

                // Save the changes made to the workbook
                excelWorkbook.Save();

                //Close the Excel workbook
                excelWorkbook.Close();

                //Release the COM objects to free up resources
                Marshal.ReleaseComObject(worksheets);

                //Quit and close the Excel application
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                // If an exception occurs, write a log entry with error details
                Task.Factory.StartNew(() => WriteLog(" ", "", ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

            }
        }

        public static void ExcelUpdateActionSheetOne(string workbookPath)
        {
            try
            {
                // Create a new instance of Microsoft Excel application
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application
                {
                    DisplayAlerts = false
                };

                // Open the specified Excel workbook
                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                // Get a reference to the worksheets in the workbook
                Microsoft.Office.Interop.Excel.Sheets worksheets = excelWorkbook.Worksheets;



                // Delete the worksheet named "Sheet1"
                excelWorkbook.Sheets["Sheet1"].Delete();

                // Delete the worksheet named "Evaluation Warning"
                excelWorkbook.Sheets["Evaluation Warning"].Delete();








                // Save the changes made to the workbook
                excelWorkbook.Save();

                //Close the Excel workbook
                excelWorkbook.Close();

                //Release the COM objects to free up resources
                Marshal.ReleaseComObject(worksheets);

                //Quit and close the Excel application
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                // If an exception occurs, write a log entry with error details
                Task.Factory.StartNew(() => WriteLog(" ", "", ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

            }
        }

        public static void SplitWrkBk(string sourceFilePath, string sheetName, string savePath)
        {
            try
            {
                // Load Excel file
                string[] files = Directory.GetFiles(sourceFilePath);
                var wb = new Aspose.Cells.Workbook(files[0]);
                // Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook(sourceFilePath);

                // Get all worksheets
                WorksheetCollection collection = wb.Worksheets;

                // Format today's date to match how it's likely in the sheet name
                // string todayStr = DateTime.Now.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture); // e.g., 17-Apr-2025

                // Find the first worksheet whose name contains today's date
                Aspose.Cells.Worksheet targetSheet = null;
                foreach (Aspose.Cells.Worksheet ws in collection)
                {
                    if (ws.Name.Contains(sheetName))
                    {
                        targetSheet = ws;
                        break;
                    }
                }

                if (targetSheet != null)
                {
                    // Create a new workbook and copy the target sheet into it
                    Aspose.Cells.Workbook wb1 = new Aspose.Cells.Workbook();
                    wb1.Worksheets[0].Copy(targetSheet);

                    // Save the new workbook with the target sheet name
                    string fileName = targetSheet.Name + ".xlsx";
                    string newWorkbookName = Path.Combine(savePath, fileName);
                    wb1.Save(newWorkbookName);

                    ExcelUpdateAction(newWorkbookName);
                }
                else
                {
                    Console.WriteLine($"No ({sheetName}) sheet found.");
                }
            }
            catch (Exception ex)
            {
                Task.Factory.StartNew(() => WriteLog(" ", " ", ex.Message + "  || " + ex.StackTrace,
                    ConfigurationManager.AppSettings["ApplicationName"],
                    string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
            }
        }

        public static bool ReadAllFiles(string sourcePath, out List<string> filePath, string extension)
        {
            bool worked = false;
            filePath = new List<string>();

            try
            {
                int counter = 0;

                // Enumerate all files in the specified source directory with the given file extension
                foreach (string file in Directory.EnumerateFiles(sourcePath, "*." + extension))
                {
                    //Add the file path to the list
                    filePath.Add(file);
                    counter++;
                }

                //If at least one file was found
                if (counter > 0)
                {
                    // Remove any null or empty file paths from the list (if any)
                    filePath.RemoveAll(item => string.IsNullOrEmpty(item));
                    worked = true;
                }
            }
            catch (Exception ex)
            {
                // If an exception occurs, write a log entry with error details
                Task.Factory.StartNew(() => WriteLog(" ", sourcePath, ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
            }

            //  // Return a boolean indicating whether the operation worked
            return worked;
        }

        public static string ReadExcelToJson(string inputPath, string destination, string fileName)
        {
            string jsonInput = string.Empty;

            try
            {
                // Create a new instance of the Aspose.Cells.Workbook class using the provided input path
                var workbook = new Aspose.Cells.Workbook(inputPath);

                // Define the JSON file path where the Excel data will be saved in JSON format
                string jsonPath = destination + fileName + ".json";

                // Save the contents of the Excel workbook as a JSON file
                workbook.Save(jsonPath);

                // Dispose of the workbook object to release resources
                workbook.Dispose();

                // Read the contents of the JSON file back into a string
                jsonInput = File.ReadAllText(jsonPath);

            }
            catch (Exception ex)
            {
                Task.Factory.StartNew(() => WriteLog("", fileName, ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
            }
            // Return the JSON data as a string
            return jsonInput;
        }

        public static void ReadJson(string jsonInput, out List<GhLinkReport> report)
        {
            // Initialize the report as an empty list of ExportWorksheet objects
            report = new List<GhLinkReport>();
            try
            {
                // Deserialize the JSON input string into a list of ExportWorksheet objects
                report = JsonConvert.DeserializeObject<List<GhLinkReport>>(jsonInput);
            }
            catch (Exception ex)
            {
                // If an exception occurs during deserialization, write a log entry with error details
                Task.Factory.StartNew(() => WriteLog(" ", " ", ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

            }
        }

        public static void ReadJsonOne(string jsonInput, out List<OpenItems> report)
        {
            // Initialize the report as an empty list of ExportWorksheet objects
            report = new List<OpenItems>();
            try
            {
                // Deserialize the JSON input string into a list of ExportWorksheet objects
                report = JsonConvert.DeserializeObject<List<OpenItems>>(jsonInput);
            }
            catch (Exception ex)
            {
                // If an exception occurs during deserialization, write a log entry with error details
                Task.Factory.StartNew(() => WriteLog(" ", " ", ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

            }
        }

        public static string ReconsKeyFromAddlText(string metaData)
        {
            string reconsKey = string.Empty;
            string result = string.Empty;
            char A = 'Â';

            try
            {
                //string inputString = "DD 26.03.24 0150414493400 2030192321519 300302 408613378330";

                // Split the string by space
                string[] parts = metaData.Trim().Split(' ', '\t');

                // Access the last element of the array
                result = parts[parts.Length - 1];

                result = Regex.Match(result, @"\d+").Value;


                reconsKey = result;




            }
            catch (Exception ex)
            {
                //Task.Factory.StartNew(() => WriteLog(" ", " ", ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

            }
            return reconsKey;
        }

        public static string ReconsKeyFromAddlTextATMUSER(string input)
        {
            string reconsKey = string.Empty;
            string result = string.Empty;
            char A = 'Â';

            try
            {

                Match match = Regex.Match(input, @"\b\d{12}\b");

                if (match.Success)
                {
                    result = match.Value;
                }


                reconsKey = result;

            }
            catch (Exception ex)
            {
                //Task.Factory.StartNew(() => WriteLog(" ", " ", ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

            }
            return reconsKey;
        }

        public static string ReconsKeyFromAddlTextATMUSEROld(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return null;

            string[] parts = input.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            string lastNumeric = null;

            foreach (string part in parts)
            {
                if (Regex.IsMatch(part, @"^[A-Za-z]")) // stop when first text appears
                    break;

                if (Regex.IsMatch(part, @"^\d+$")) // purely numeric
                    lastNumeric = part;
            }
            return lastNumeric;
        }

        public static bool GetOpenItemsFile(string openItemsPath, string generatedJson, string source, out List<OpenItems> openItems, out string message)
        {
            openItems = new List<OpenItems>();
            bool success = false;
            message = "Unable to get open items file";

            try
            {
                string fileNam = string.Empty;
                //Delete files in source folder
                string[] sourcefolder = Directory.GetFiles(source);
                foreach (var sourcefile in sourcefolder)
                {
                    File.Delete(sourcefile);
                }

                UserFunctions.SplitWrkBk(openItemsPath, StaticVariables.OPENITEMS, source);
                if (!UserFunctions.ReadAllFiles(source, out List<string> filePath1, "xlsx"))
                {
                    Console.WriteLine("No data found in source Path");

                    Task.Factory.StartNew(() => UserFunctions.WriteLog(" ", "", "No data found in source Path", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
                }

                Console.WriteLine(" ");

                Console.WriteLine("Data found in specified location");

                Task.Factory.StartNew(() => UserFunctions.WriteLog(" ", "", "Data found in specified location", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

                foreach (var item1 in filePath1)
                {
                    string fileName = Path.GetFileNameWithoutExtension(item1);
                    string jsonInput = UserFunctions.ReadExcelToJson(item1, generatedJson, fileName);

                    //string message = string.Empty;

                    if (string.IsNullOrEmpty(jsonInput))
                    {
                        Console.WriteLine("Unable to read data from " + item1);

                        Task.Factory.StartNew(() => UserFunctions.WriteLog(" ", fileName, "Data read from " + item1 + " successfully", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));



                        Thread.Sleep(1500);

                        continue;
                    }

                    Console.WriteLine(" ");
                    Console.WriteLine("Data read from " + item1 + " successfully");

                    Task.Factory.StartNew(() => UserFunctions.WriteLog(" ", fileName, "Data read from " + item1 + " successfully", ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));


                    UserFunctions.ReadJsonOne(jsonInput, out openItems);


                }

                

                //Delete files in source folder
                sourcefolder = Directory.GetFiles(source);
                foreach (var sourcefile in sourcefolder)
                {
                    File.Delete(sourcefile);
                }


                success = true;
                message = "Open items gotten successfully";
            }
            catch (Exception ex)
            {
                Task.Factory.StartNew(() => WriteLog(" ", " ", ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

            }
            return success;

        }

        public static bool MergeOpenItemsNReport(List<GhLinkReport> ghLinkReport, List<OpenItems> openItems, out List<GhLinkReport> newGhLinkReport, out string message)
        {
            List<GhLinkReport> openItemsExceptions = new List<GhLinkReport>();
            List<List<GhLinkReport>> reportOpenitems = new List<List<GhLinkReport>>();
            newGhLinkReport = new List<GhLinkReport>();
            bool success = false;
            message = "Unable to merge carry forward data and report";

            try
            {
                if (openItems.Any())
                {
                    openItemsExceptions = openItems.Select(x => new GhLinkReport
                    {
                        LCY_AMT = x.GrandTotal,
                        VALUE_DT = x.VALUE_DT,
                        ADDL_TEXT = x.ADDL_TEXT,
                        USER_ID = x.USER_ID,


                    }).ToList();
                }

                reportOpenitems.Add(ghLinkReport);
                reportOpenitems.Add(openItemsExceptions);

                //report.Clear();

                foreach (var item in reportOpenitems)
                {
                    newGhLinkReport.AddRange(item);
                }



                success = true;
                message = "Carry forward and report merge was successful";


            }
            catch (Exception ex)
            {
                Task.Factory.StartNew(() => WriteLog(" ", " ", ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));

            }
            return success;
        }

        public static bool GetReconKey(List<GhLinkReport> ghLinkReports, out List<GhLinkInd> ghLinkInd, out string message)
        {
            ghLinkInd = new List<GhLinkInd>();

            bool success = false;
            message = "Unable to get recon key";

            try
            {
                ghLinkReports = ghLinkReports.Where(x => x != null).ToList();
                ghLinkReports = ghLinkReports.Where(x => !string.IsNullOrWhiteSpace(x.USER_ID)).ToList();

                ghLinkReports.ForEach(x =>
                {
                    if (!string.IsNullOrWhiteSpace(x.USER_ID) && !x.USER_ID.Contains(StaticVariables.USER))
                    {
                        x.ATM_RRN = ReconsKeyFromAddlText(x.ADDL_TEXT);
                    }
                    else
                    {
                        x.ATM_RRN = ReconsKeyFromAddlTextATMUSER(x.ADDL_TEXT);
                    }
                });


                if (ghLinkReports.Any())
                {
                    ghLinkInd = ghLinkReports.Select(x => new GhLinkInd
                    {
                        VALUE_DT = x.VALUE_DT,
                        ADDL_TEXT = x.ADDL_TEXT,
                        ATM_RRN = x.ATM_RRN,
                        LCY_AMT = x.LCY_AMT,
                        USER_ID = x.USER_ID,

                    }).ToList();
                }

                ghLinkInd = ghLinkInd.Where(x => !string.IsNullOrWhiteSpace(x.USER_ID)).ToList();

                ghLinkInd.ForEach(x =>
                {
                    if (x.LCY_AMT.StartsWith(StaticVariables.NEGATIVE))
                    {
                        x.DC = StaticVariables.D;
                    }
                    else
                    {
                        x.DC = StaticVariables.C;
                    }

                });

                success = true;
                message = "Recon key generated successfully";

            }
            catch (Exception ex)
            {
                // Log any exceptions that occur during execution
                Task.Factory.StartNew(() => WriteLog(" ", " ", ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
            }

            return success; // Return the operation status
        }

        public static bool WriteToExcelSheet(List<string> jsonInput, string excelPath, string savePath, string sheetName, string filename, out string message)
        {
            bool success = false;
            message = "Unable to write to sheet";

            try
            {
                var workbook = new Aspose.Cells.Workbook(excelPath); // Load the Excel workbook from the specified path
                Aspose.Cells.Worksheet worksheet = workbook.Worksheets[sheetName]; // Access the specified worksheet by name

                worksheet.Cells.Clear();






                // Find the last row index with data in the worksheet
                int lastRowIndex = worksheet.Cells.MaxDataRow + 1;


                foreach (var item in jsonInput)
                {
                    // Set JsonLayoutOptions to configure JSON import settings
                    JsonLayoutOptions options = new JsonLayoutOptions
                    {
                        ArrayAsTable = true // Import arrays as tables
                    };

                    // Import JSON Data into the worksheet starting from the last row index
                    JsonUtility.ImportData(item, worksheet.Cells, lastRowIndex, 0, options);
                }

                // Generate a filename with the current date and time
                //string currentDateAndTime = DateTime.Now.ToString("dd-MM-yyyy");
                string fileName = filename + "." + "xlsx";
                string save = savePath + "\\" + fileName;

                // Save the modified workbook to the specified path
                workbook.Save(save);
                workbook.Dispose(); // Dispose the workbook to release resources
                ExcelUpdateAction(save); // Perform an action after updating the Excel file

                success = true; // Operation completed successfully
                message = "Write to sheet was successful"; // Update the success message
            }
            catch (Exception ex)
            {
                // Log any exceptions that occur during execution
                Task.Factory.StartNew(() => WriteLog(" ", " ", ex.Message + "  || " + ex.StackTrace, ConfigurationManager.AppSettings["ApplicationName"], string.Format("{0}.{1}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name)));
            }

            return success; // Return the operation status
        }
    }
}
