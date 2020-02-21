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
using System.Linq;
using ExcelDataReader;
using System.Data;
using System.Data.SqlClient;

namespace AgentConsoleApp
{
    class Program
    {
        public class returnModel
        {
            public string FileName { get; set; }
            public int RowNo { get; set; }
        }

        static void Main(string[] args)
        {
            var conn = ConfigurationManager.AppSettings["DBConnectionString"].ToString();
            string sourceDirectory;
            List<returnModel> returnCollection = new List<returnModel>();
            string fileName = "";
            int counterFile = 1;
            int counterLine;
            int counterFileValidate = 1;

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
            int LineNum;
            int ColumnNum;
            string sheetName = "Sheet1";
            string TableName = "[SAPMM-WM]";
            int i = 0;
            int ColumnNumChecker = 53;
            bool sheetChecker = true;

            try
            {
                // Full path for txt
                var FilePath = Directory.EnumerateFiles(sourceDirectory, "*.*", SearchOption.TopDirectoryOnly).Where(s => s.ToLower().EndsWith(".xls") || s.ToLower().EndsWith(".xlsx"));

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
                                    while (reader.Read())
                                    {
                                        if (i > 0)
                                        {
                                            int rowNO = i + 1;

                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 1, 10, "B", "Material Type");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 2, 255, "C", "Material Type Description");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 3, 10, "D","Material Group");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 4, 255, "E", "Matl Grp Desc#");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 5, 12, "F", "Material");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 6, 255, "G", "Description");
                                            ValidateDate(reader, pbValidate, rowNO, counterFileValidate, 7, "H", "Posting Date");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 8, 255, "I", "ได้รับมาจาก");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 9, 255, "J", "จ่ายไปให้");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 10, 10, "K", "Movement type");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 11, 255, "L", "Mvt Type Text");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 12, 100, "M", "Batch");
                                            ValidateDate(reader, pbValidate, rowNO, counterFileValidate, 13, "N", "MFG Date");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 14, 100, "O", "Manufacturer Batch");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 15, 100, "P", "Manufacturer");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 16, 255, "Q", "Manufacturer Name");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 17, 100, "R", "Vendor");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 18, 255, "S", "Vendor Name");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 19, 20, "T", "Sold-to");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 20, 255, "U", "Sold-to Name");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 21, 255, "V", "Sold-to Address");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 22, 100, "W", "Sold-to Province");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 23, 20, "X", "Ship-to");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 24, 255, "Y", "Ship-to Name");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 25, 255, "Z", "Ship-to Address");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 26, 100, "AA", "Ship-to Province");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 27, 5, "AB", "Customer Group 1");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 28, 100, "AC", "Customer Group 1 - Desc#");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 29, 5, "AD", "Customer Group 2");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 30, 100, "AE", "Customer Group 2 - Desc#");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 31, 5, "AF", "Customer Group 3");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 32, 100, "AG", "Customer Group 3 - Desc#");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 33, 20, "AH", "FG material");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 34, 255, "AI", "FG Material Description");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 35, 100, "AJ", "FG Batch");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 36, 5, "AK", "Cost Center");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 37, 255, "AL", "Cost Center Description");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 38, 10, "AM", "Plant");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 39, 10, "AN", "Storage Loc#");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 40, 10, "AO", "Dest# Plant");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 41, 10, "AP", "Dest# Sloc");
                                            ValidateFloat(reader, pbValidate, rowNO, counterFileValidate, 42, "AQ", "ยอดยกมา");
                                            ValidateFloat(reader, pbValidate, rowNO, counterFileValidate, 43, "AR", "ปริมาณรับ");
                                            ValidateFloat(reader, pbValidate, rowNO, counterFileValidate, 44, "AS", "ปริมาณจ่าย");
                                            ValidateFloat(reader, pbValidate, rowNO, counterFileValidate, 45, "AT", "ปริมาณคงเหลือ");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 46, 20, "AU", "Unit");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 47, 255, "AV", "หมายเหตุ");
                                            ValidateDate(reader, pbValidate, rowNO, counterFileValidate, 48, "AW", "Entered on");
                                            ValidateDate(reader, pbValidate, rowNO, counterFileValidate, 49, "AX", "Entered at");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 50, 20,"AY", "Material Doc#");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 51, 10, "BZ", "Mat# Doc# Year");
                                            ValidateString(reader, pbValidate, rowNO, counterFileValidate, 52, 5, "BA", "Mat# Doc#Item");
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
                    LineNum = 0;
                    ColumnNum = 0;
                    counterLine = 1;

                    returnModel Model = new returnModel();

                    fileName = Path.GetFileName(currentFile);

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

                                    // Row count
                                    LineNum = dt.Rows.Count;

                                    // Column count
                                    ColumnNum = dt.Columns.Count;

                                    // Validate excel column
                                    if (ColumnNum != ColumnNumChecker)
                                    {
                                        pbOverall.Refresh(counterFile, "Import error occured");
                                        throw new ArgumentException($"Excel file must have {ColumnNumChecker} column!");
                                    }

                                    // Sanitize data
                                    foreach (DataColumn c in dt.Columns)
                                    {
                                        if (c.DataType == typeof(string))
                                        {
                                            foreach (DataRow r in dt.Rows)
                                            {
                                                r[c.ColumnName] = r[c.ColumnName].ToString().Trim();

                                                // Convert empty string into NULL
                                                if (r[c.ColumnName].ToString().Length == 0)
                                                {
                                                    r[c.ColumnName] = DBNull.Value;
                                                }
                                            }
                                        }
                                    }

                                    using (SqlBulkCopy bc = new SqlBulkCopy(conn, SqlBulkCopyOptions.UseInternalTransaction | SqlBulkCopyOptions.TableLock))
                                    {
                                        bc.DestinationTableName = TableName;
                                        bc.BatchSize = reader.RowCount;
                                        bc.ColumnMappings.Add(1, "[Material Type]");
                                        bc.ColumnMappings.Add(2, "[Material Type Description]");
                                        bc.ColumnMappings.Add(3, "[Material Group]");
                                        bc.ColumnMappings.Add(4, "[Matl Grp Desc#]");
                                        bc.ColumnMappings.Add(5, "[Material]");
                                        bc.ColumnMappings.Add(6, "[Description]");
                                        bc.ColumnMappings.Add(7, "[Posting Date]");
                                        bc.ColumnMappings.Add(8, "[ได้รับมาจาก]");
                                        bc.ColumnMappings.Add(9, "[จ่ายไปให้]");
                                        bc.ColumnMappings.Add(10, "[Movement type]");
                                        bc.ColumnMappings.Add(11, "[Mvt Type Text]");
                                        bc.ColumnMappings.Add(12, "[Batch]");
                                        bc.ColumnMappings.Add(13, "[MFG Date]");
                                        bc.ColumnMappings.Add(14, "[Manufacturer Batch]");
                                        bc.ColumnMappings.Add(15, "[Manufacturer]");
                                        bc.ColumnMappings.Add(16, "[Manufacturer Name]");
                                        bc.ColumnMappings.Add(17, "[Vendor]");
                                        bc.ColumnMappings.Add(18, "[Vendor Name]");
                                        bc.ColumnMappings.Add(19, "[Sold-to]");
                                        bc.ColumnMappings.Add(20, "[Sold-to Name]");
                                        bc.ColumnMappings.Add(21, "[Sold-to Address]");
                                        bc.ColumnMappings.Add(22, "[Sold-to Province]");
                                        bc.ColumnMappings.Add(23, "[Ship-to]");
                                        bc.ColumnMappings.Add(24, "[Ship-to Name]");
                                        bc.ColumnMappings.Add(25, "[Ship-to Address]");
                                        bc.ColumnMappings.Add(26, "[Ship-to Province]");
                                        bc.ColumnMappings.Add(27, "[Customer Group 1]");
                                        bc.ColumnMappings.Add(28, "[Customer Group 1 - Desc#]");
                                        bc.ColumnMappings.Add(29, "[Customer Group 2]");
                                        bc.ColumnMappings.Add(30, "[Customer Group 2 - Desc#]");
                                        bc.ColumnMappings.Add(31, "[Customer Group 3]");
                                        bc.ColumnMappings.Add(32, "[Customer Group 3 - Desc#]");
                                        bc.ColumnMappings.Add(33, "[FG material]");
                                        bc.ColumnMappings.Add(34, "[FG Material Description]");
                                        bc.ColumnMappings.Add(35, "[FG Batch]");
                                        bc.ColumnMappings.Add(36, "[Cost Center]");
                                        bc.ColumnMappings.Add(37, "[Cost Center Description]");
                                        bc.ColumnMappings.Add(38, "[Plant]");
                                        bc.ColumnMappings.Add(39, "[Storage Loc#]");
                                        bc.ColumnMappings.Add(40, "[Dest# Plant]");
                                        bc.ColumnMappings.Add(41, "[Dest# Sloc]");
                                        bc.ColumnMappings.Add(42, "[ยอดยกมา]");
                                        bc.ColumnMappings.Add(43, "[ปริมาณรับ]");
                                        bc.ColumnMappings.Add(44, "[ปริมาณจ่าย]");
                                        bc.ColumnMappings.Add(45, "[ปริมาณคงเหลือ]");
                                        bc.ColumnMappings.Add(46, "[Unit]");
                                        bc.ColumnMappings.Add(47, "[หมายเหตุ]");
                                        bc.ColumnMappings.Add(48, "[Entered on]");
                                        bc.ColumnMappings.Add(49, "[Entered at]");
                                        bc.ColumnMappings.Add(50, "[Material Doc#]");
                                        bc.ColumnMappings.Add(51, "[Mat# Doc# Year]");
                                        bc.ColumnMappings.Add(52, "[Mat# Doc#Item]");
                                        bc.WriteToServer(dt);
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
                        Model.RowNo = LineNum;
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

        // Method Validate Float
        public static void ValidateFloat(IExcelDataReader reader, ProgressBar pb, int rowNum, int fileCount, int columnNum, string columnXlsName, string columnName)
        {
            double columnFloat;

            // Cast type
            try
            {
                columnFloat = Convert.ToDouble(reader.GetValue(columnNum));
            }
            catch (Exception ex)
            {
                pb.Refresh(fileCount, "Validate failed.");
                throw new ArgumentException($"{ex.Message}  At Column: {columnXlsName} ({columnName}), Row: {rowNum}");
            }
        }

        // Method Validate Date
        public static void ValidateDate(IExcelDataReader reader, ProgressBar pb, int rowNum, int fileCount, int columnNum, string columnXlsName, string columnName)
        {
            DateTime columnDate;

            // Cast type
            try
            {
                columnDate = Convert.ToDateTime(reader.GetValue(columnNum));
            }
            catch (Exception ex)
            {
                pb.Refresh(fileCount, "Validate failed.");
                throw new ArgumentException($"{ex.Message}  At Column: {columnXlsName} ({columnName}), Row: {rowNum}");
            }
        }

        // Method Validate String
        public static void ValidateString(IExcelDataReader reader, ProgressBar pb, int rowNum, int fileCount, int columnNum, int strLength, string columnXlsName, string columnName)
        {
            string columnStr;

            // Cast type
            try
            {
                columnStr = Convert.ToString(reader.GetValue(columnNum));
            }
            catch (Exception ex)
            {
                pb.Refresh(fileCount, "Validate failed.");
                throw new ArgumentException($"{ex.Message}  At Column: {columnXlsName} ({columnName}), Row: {rowNum}");
            }

            // Validate length
            if (columnStr.Length > strLength)
            {
                pb.Refresh(fileCount, "Validate failed.");
                throw new ArgumentException($"The field cannot contain more than {strLength} characters. At Column : {columnXlsName} ({columnName}), Row : {rowNum}");
            }

        }

    }

}
