using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelWorkbookRefresher
{
    class Program
    {
        public static Application _currentApplication;

        public static void Main(string[] args)
        {
            // write progress
            WriteStartProgress("Connecting to Excel...");

            _currentApplication = new Application();

            // write progress
            WriteEndProgress("Connected.");

            Console.Write("Enter directory to scan (or press enter for current directory): ");

            string folderPath = Console.ReadLine();
            if (String.IsNullOrWhiteSpace(folderPath) == true)
            {
                folderPath = Directory.GetCurrentDirectory();
            }

            Console.WriteLine();
            
            // write progress
            WriteStartProgress("Executing refresh operation...");

            Console.WriteLine();

            try
            {
                RefreshData(folderPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.Message.ToString());
                
                if (_currentApplication != null)
                {
                    _currentApplication.Workbooks.Close();
                    _currentApplication.Quit();
                }
            }

            // write progress
            WriteEndProgress("Operation completed.");

            Console.WriteLine();
            Console.Write("Press any key to continue: ");

            Console.ReadLine();
        }

        private static void RefreshData(string directory)
        {
            if (directory.Substring(directory.Length - 1, 1) != @"\")
            {
                directory = directory + @"\";
            }

            DirectoryInfo files = new DirectoryInfo(directory);

            foreach (FileInfo fileName in files.GetFiles("*.xls*"))
            {
                if (fileName.ToString().Substring(0, 2) == "~$") continue;

                // write progress
                WriteStartProgress("Processing '" + fileName.ToString() + "'...");

                if (IsFileLocked(fileName) == false)
                {
                    try
                    {
                        Workbook wb = _currentApplication.Workbooks.Open(fileName.FullName, false, false, Type.Missing, "", "", true, XlPlatform.xlWindows, "", false, false, 0, false, true, 0);

                        foreach (WorkbookConnection wc in wb.Connections)
                        {
                            wc.TextConnection.TextFilePromptOnRefresh = false;
                        }

                        wb.RefreshAll();
                        wb.Save();

                        _currentApplication.Workbooks.Close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("ERROR: Could not refresh '" + fileName.ToString() + "'. " + ex.Message.ToString());
                    }

                    // write progress
                    WriteEndProgress("Complete.");
                }
                else
                {
                    // write progress
                    WriteEndProgress("ERROR: File is locked or in use.");
                }
            }
        }

        private static bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            return false;
        }

        private static void WriteStartProgress(string progress)
        {
            Console.WriteLine(progress);
        }

        private static void WriteEndProgress(string result)
        {
            Console.WriteLine(result);
            Console.WriteLine();
        }
    }
}
