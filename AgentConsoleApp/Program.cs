using ConsoleTables;
using Figgle;
using System;
using System.Drawing;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Threading;
using Konsole;
using Console = Colorful.Console;
using static AgentConsoleApp.ImportController;
using System.Linq;
using ExcelDataReader;
using System.Data;

namespace AgentConsoleApp
{
    class Program
    {
        public class returnModel
        {
            public string FileName { get; set; }
            public int HeaderNo { get; set; }
            public int DetailNo { get; set; }
        }

        static void Main(string[] args)
        {
            var conn = ConfigurationManager.AppSettings["DBConnectionString"].ToString();
            string sourceDirectory;
            string line;
            int detailLineNo, headerLineNo;
            List<returnModel> returnCollection = new List<returnModel>();
            string FL_Filecode;
            string FL_TotalRecord;
            string HD_PoNo;
            string fileName = "";
            int counterFile = 1;
            int counterFileValidate = 1;
            int counterLine;

            #region Fancy header
            /*
            Console.Write(FiggleFonts.Ogre.Render("------------"));
            List<char> chars = new List<char>()
            {
                ' ', 'C', 'r', 'e', 'a', 't', 'e', 'd', ' ',
                'b', 'y', ' ',
                'P', 'i', 'r', 'i', 'y', 'a', 'V', ' '
            };
            Console.Write("---------------", Color.LawnGreen);
            Console.WriteWithGradient(chars, Color.Blue, Color.Purple, 16);
            Console.Write("---------------", Color.LawnGreen);
            Console.WriteLine("\n");
            */
            #endregion

            // Ask the user to type path
            if (args.Length == 0)
            {
                // Display title
                Console.Title = "ExcelToDB 1.00";

                // Display header
                Console.WriteWithGradient(FiggleFonts.Banner.Render("excel to db"), Color.LightGreen, Color.ForestGreen, 16);
                Console.ReplaceAllColorsWithDefaults();

                // Display copyright
                Console.WriteLine(" ---------------------- Created by PiriyaV -----------------------\n", Color.LawnGreen);

                Console.Write(@"Enter source path (eg: D:\folder) : ", Color.LightYellow);
                sourceDirectory = Convert.ToString(Console.ReadLine());
                Console.Write("\n");
            }
            else
            {
                sourceDirectory = Convert.ToString(args[0]);
            }

            // Variable for backup
            string folderBackup = "imported_" + DateTime.Now.ToString("ddMMyyyy_HHmmss");
            string folderBackupPath = Path.Combine(sourceDirectory, folderBackup);

            // Initial values
            int i = 0;
            bool sheetChecker = true;
            string sheetName = "Sheet1";
            string TableName = "SAPMM-WM";

            try
            {
                // Full path for txt
                var FilePath = Directory.EnumerateFiles(sourceDirectory, "*.*", SearchOption.AllDirectories).Where(s => s.ToLower().EndsWith(".xls") || s.ToLower().EndsWith(".xlsx"));

                // Count txt file
                DirectoryInfo di = new DirectoryInfo(sourceDirectory);
                int FileNumXls = di.GetFiles("*.xls").Length;
                int FileNumXlsx = di.GetFiles("*.xlsx").Length;
                int FileNum = FileNumXls + FileNumXlsx;

                // Throw no txt file
                if (FileNum == 0)
                {
                    throw new ArgumentException("Excel file not found in folder.");
                }

                #region Validate Section
                var pbValidate = new ProgressBar(PbStyle.DoubleLine, FileNum);

                foreach (string currentFile in FilePath)
                {
                    sheetChecker = true;

                    // Update progress bar (Overall)
                    fileName = Path.GetFileName(currentFile);
                    pbValidate.Refresh(counterFileValidate, "Validating, Please wait...");
                    Thread.Sleep(50);

                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                    using (var stream = File.Open(currentFile, FileMode.Open, FileAccess.Read))
                    {
                        // ExcelDataReader Config
                        var conf = new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        };

                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            // Validation Excel file
                            do
                            {
                                if (reader.Name == sheetName)
                                {
                                    sheetChecker = false;
                                    string columnStr;
                                    DateTime columnDateTime;
                                    while (reader.Read())
                                    {
                                        if (i > 0)
                                        {
                                            int rowNO = i + 1;

                                            #region Material Type
                                            // Cast type
                                            try
                                            {
                                                columnStr = Convert.ToString(reader.GetValue(1));
                                            }
                                            catch (Exception ex)
                                            {
                                                pbValidate.Refresh(counterFileValidate, "Validate failed.");
                                                throw new ArgumentException($"{ex.Message}  At Column: B (Material Type), Row: {rowNO}");
                                            }

                                            // Validate length
                                            if (columnStr.Length > 10)
                                            {
                                                pbValidate.Refresh(counterFileValidate, "Validate failed.");
                                                throw new ArgumentException($"The field cannot contain more than 10 characters. At Column : B (Material Type), Row : {rowNO}");
                                            }
                                            #endregion

                                            #region Material Type Description
                                            // Cast type
                                            try
                                            {
                                                columnStr = Convert.ToString(reader.GetValue(2));
                                            }
                                            catch (Exception ex)
                                            {
                                                pbValidate.Refresh(counterFileValidate, "Validate failed.");
                                                throw new ArgumentException($"{ex.Message}  At Column: C (Material Type Description), Row: {rowNO}");
                                            }

                                            // Validate length
                                            if (columnStr.Length > 255)
                                            {
                                                pbValidate.Refresh(counterFileValidate, "Validate failed.");
                                                throw new ArgumentException($"The field cannot contain more than 255 characters. At Column : C (Material Type Description), Row : {rowNO}");
                                            }
                                            #endregion

                                            #region Material Group
                                            // Cast type
                                            try
                                            {
                                                columnStr = Convert.ToString(reader.GetValue(3));
                                            }
                                            catch (Exception ex)
                                            {
                                                pbValidate.Refresh(counterFileValidate, "Validate failed.");
                                                throw new ArgumentException($"{ex.Message}  At Column: D (Material Group), Row: {rowNO}");
                                            }

                                            // Validate length
                                            if (columnStr.Length > 10)
                                            {
                                                pbValidate.Refresh(counterFileValidate, "Validate failed.");
                                                throw new ArgumentException($"The field cannot contain more than 10 characters. At Column : D (Material Group), Row : {rowNO}");
                                            }
                                            #endregion

                                            #region Matl Grp Desc.
                                            // Cast type
                                            try
                                            {
                                                columnStr = Convert.ToString(reader.GetValue(4));
                                            }
                                            catch (Exception ex)
                                            {
                                                pbValidate.Refresh(counterFileValidate, "Validate failed.");
                                                throw new ArgumentException($"{ex.Message}  At Column: E (Matl Grp Desc.), Row: {rowNO}");
                                            }

                                            // Validate length
                                            if (columnStr.Length > 255)
                                            {
                                                pbValidate.Refresh(counterFileValidate, "Validate failed.");
                                                throw new ArgumentException($"The field cannot contain more than 255 characters. At Column : E (Matl Grp Desc.), Row : {rowNO}");
                                            }
                                            #endregion
                                        }
                                    }

                                }
                            } while (reader.NextResult());
                        }
                    }

                    // Change wording in progress bar
                    if (counterFileValidate == FileNum)
                    {
                        pbValidate.Refresh(counterFileValidate, "Validate finished.");
                    }

                    // Return error if excel file don't have specific sheet name 
                    if (sheetChecker)
                    {
                        pbValidate.Refresh(counterFileValidate, "Validate failed.");
                        throw new ArgumentException($"Excel must have sheet name \" {sheetName} \"");
                    }

                    counterFileValidate++;
                }
                #endregion

                #region Import Section
                // Create progress bar (Overall)
                var pbOverall = new ProgressBar(PbStyle.DoubleLine, FileNum);

                foreach (string currentFile in FilePath)
                {
                    // Initial variable
                    headerLineNo = 0;
                    detailLineNo = 0;
                    counterLine = 1;
                    FL_Filecode = "";
                    FL_TotalRecord = "";
                    HD_PoNo = "";

                    returnModel Model = new returnModel();

                    fileName = Path.GetFileName(currentFile);

                    // Create progress bar (Each file)
                    int LineNum = CountLinesReader(currentFile);
                    var pbDetail = new ProgressBar(PbStyle.SingleLine, LineNum);

                    // Update progress bar (Overall)
                    pbOverall.Refresh(counterFile, "Importing, Please wait...");
                    Thread.Sleep(50);

                    using (var stream = File.Open(currentFile, FileMode.Open, FileAccess.Read))
                    {
                        // ExcelDataReader Config
                        var conf = new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        };

                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            // Validation Excel file
                            do
                            {
                                if (reader.Name == sheetName)
                                {
                                    // Read as DataSet
                                    var result = reader.AsDataSet(conf);

                                    // Convert to Datatable
                                    DataTable dt = result.Tables[sheetName];

                                    // Sanitize data
                                    foreach (DataColumn c in dt.Columns)
                                    {
                                        if (c.DataType == typeof(string))
                                        {
                                            foreach (DataRow r in dt.Rows)
                                            {
                                                //try
                                                //{
                                                r[c.ColumnName] = r[c.ColumnName].ToString().Trim();
                                                //}
                                                //catch
                                                //{ }
                                            }
                                        }
                                    }
                                }
                            } while (reader.NextResult());
                        counterLine++;
                    }

                    // Create folder for file import successful
                    if (!Directory.Exists(folderBackupPath))
                    {
                        Directory.CreateDirectory(folderBackupPath);
                    }

                    // Move file to folder backup
                    string destFile = Path.Combine(folderBackupPath, fileName);
                    File.Move(currentFile, destFile);

                    // Add detail to model for showing in table
                    Model.HeaderNo = headerLineNo;
                    Model.DetailNo = detailLineNo;
                    Model.FileName = fileName;
                    returnCollection.Add(Model);

                    // Change wording in progress bar
                    if (counterFile == FileNum)
                    {
                        pbOverall.Refresh(counterFile, "Import finished.");
                    }

                    counterFile++;
                }
                #endregion
            }
            catch (Exception ex)
            {
                //pbOverall.Refresh(counterFile, "Import failed");

                // Show error message
                Console.Write("\nError occured : " , Color.OrangeRed);
                Console.WriteLine(ex.Message);
                //Console.WriteLine("Error trace : " + ex.StackTrace);

                // Show error on
                if (!String.IsNullOrEmpty(fileName))
                {
                    Console.Write("\nError on : ", Color.OrangeRed);
                    Console.WriteLine("'" + fileName + "'");
                }
                
                // Show description
                Console.WriteLine("\nPlease check your path or file and try again.\n", Color.Yellow);
            }
            finally
            {
                // Show table
                if (returnCollection.Count > 0)
                {
                    Console.WriteLine("\n--------------- Imported detail ---------------", Color.LightGreen);
                    ConsoleTable.From(returnCollection).Write();
                }
                //Console.WriteLine(JsonSerializer.Serialize(returnCollection));

                // Show backup folder path
                if (Directory.Exists(folderBackupPath))
                {
                    Console.Write("\nImported folder : ", Color.LightGreen);
                    Console.WriteLine($"'{ folderBackupPath }'");
                }
            }

            // Wait key to terminate
            Console.Write("\nPress any key to close this window ");
            Console.ReadKey();
        }

    }

}
