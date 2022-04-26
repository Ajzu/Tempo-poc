using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.OleDb;
using DataTable = System.Data.DataTable;

namespace well_estimation_tool_grid_poc.Helpers
{
    public class ExcelHelpers
    {
        public void BeginExcelWorkFlow()
        {
            try
            {
                SaveDataSetAsExcel(CreateDataSet(), @"H:\POCTestDocuments");// @"\\BLRNCS01C02\arsing5$\Sync\POCTestDocuments"); // @"C:\Users\arsing5\Desktop\POCTestDocuments");
            }
            catch (Exception ex)
            {

            }
        }
        public DataSet CreateDataSet()
        {
            try
            {
                // Create 2 DataTable instances.
                DataTable table1 = new DataTable("patients");
                table1.Columns.Add("name");
                table1.Columns.Add("id");
                table1.Rows.Add("sam", 1);
                table1.Rows.Add("mark", 2);

                DataTable table2 = new DataTable("medications");
                table2.Columns.Add("id");
                table2.Columns.Add("medication");
                table2.Rows.Add(1, "atenolol");
                table2.Rows.Add(2, "amoxicillin");

                // Create a DataSet and put both tables in it.
                DataSet set = new DataSet("office");
                set.Tables.Add(table1);
                //set.Tables.Add(table2);//original- sending only one datatable later will send multiple.

                return set;
            }
            catch (Exception ex)
            {
                return null;
            }
        }


        public void SaveDataSetAsExcel(DataSet dataset, string excelFilePath)
        {
            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(excelFilePath, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                    DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbookPart.Workbook.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheets());

                    foreach (DataTable table in dataset.Tables)
                    {
                        UInt32Value sheetCount = 0;
                        sheetCount++;

                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                        var sheetData = new SheetData();
                        worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                        Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = sheetCount, Name = table.TableName };
                        sheets.AppendChild(sheet);

                        Row headerRow = new Row();

                        List<string> columns = new List<string>();
                        foreach (System.Data.DataColumn column in table.Columns)
                        {
                            columns.Add(column.ColumnName);

                            Cell cell = new Cell();
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(column.ColumnName);
                            headerRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(headerRow);

                        foreach (DataRow dsrow in table.Rows)
                        {
                            Row newRow = new Row();
                            foreach (String col in columns)
                            {
                                Cell cell = new Cell();
                                cell.DataType = CellValues.String;
                                //cell.CellValue = new CellValue(dsrow[col].ToString());//original
                                if (dsrow[col] != null)
                                    cell.CellValue = new CellValue(Convert.ToString(dsrow[col]));
                                newRow.AppendChild(cell);
                            }

                            sheetData.AppendChild(newRow);
                        }



                    }
                    workbookPart.Workbook.Save();
                }
            }
            catch (Exception ex)
            {

            }
        }


        public void ReadExcel_Dummy2()
        {
            try
            {
                string physicalPath = @"C:\Users\arsing5\Desktop\Test\well-estimation-tool-grid-poc\well-estimation-tool-grid-poc\test\HelloWorld.xlsx";// "Your Excel file physical path";
                OleDbCommand cmd = new OleDbCommand();
                OleDbDataAdapter da = new OleDbDataAdapter();
                DataSet ds = new DataSet();
                String strNewPath = physicalPath;
                String connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strNewPath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                String query = "SELECT * FROM [Sheet1$]"; // You can use any different queries to get the data from the excel sheet
                OleDbConnection conn = new OleDbConnection(connString);
                if (conn.State == ConnectionState.Closed) conn.Open();
                try
                {
                    cmd = new OleDbCommand(query, conn);
                    da = new OleDbDataAdapter(cmd);
                    da.Fill(ds);

                }
                catch
                {
                    // Exception Msg 

                }
                finally
                {
                    da.Dispose();
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                // throw;
            }
        }
        public void ReadExcel_Dummy1()
        {
            try
            {
                string fileLocation = @"C:\Users\arsing5\Desktop\Test\well-estimation-tool-grid-poc\well-estimation-tool-grid-poc\test\HelloWorld.xlsx"; ;
                DataTable sheet1 = new DataTable("Excel Sheet");
                OleDbConnectionStringBuilder csbuilder = new OleDbConnectionStringBuilder();
                csbuilder.Provider = "Microsoft.ACE.OLEDB.12.0";
                csbuilder.DataSource = fileLocation;
                csbuilder.Add("Extended Properties", "Excel 12.0 Xml;HDR=YES");
                string selectSql = @"SELECT * FROM [Sheet1$]";
                using (OleDbConnection connection = new OleDbConnection(csbuilder.ConnectionString))
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(selectSql, connection))
                {
                    connection.Open();
                    adapter.Fill(sheet1);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }


        //private DataTable GetDataTable(string sql, string connectionString)
        //{
        //    DataTable dt = new DataTable();

        //    using (OLEDBConnection conn = new OleDbConnection(connectionString))
        //    {
        //        conn.Open();
        //        using (OleDbCommand cmd = new OleDbCommand(sql, conn))
        //        {
        //            using (OleDbDataReader rdr = cmd.ExecuteReader())
        //            {
        //                dt.Load(rdr);
        //                return dt;
        //            }
        //        }
        //    }
        //}

        //private void GetExcel()
        //{
        //    string fullPathToExcel = "<Path to Excel file>"; //ie C:\Temp\YourExcel.xls
        //    string connString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=yes'", fullPathToExcel);
        //    DataTable dt = GetDataTable("SELECT * from [SheetName$]", connString);

        //    foreach (DataRow dr in dt.Rows)
        //    {
        //        //Do what you need to do with your data here
        //    }
        //}


        /// <summary>
        /// Link with solution - https://stackoverflow.com/questions/23041021/how-to-write-some-data-to-excel-file-xlsx
        /// </summary>
        //public void CreateExcel_Dummy1()
        //{
        //    object misvalue = System.Reflection.Missing.Value;
        //    //Start Excel and get Application object.
        //    oXL = new Microsoft.Office.Interop.Excel.Application();
        //    oXL.Visible = true;

        //    //Get a new workbook.
        //    oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
        //    oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

        //    //Add table headers going cell by cell.
        //    oSheet.Cells[1, 1] = "First Name";
        //    oSheet.Cells[1, 2] = "Last Name";
        //    oSheet.Cells[1, 3] = "Full Name";
        //    oSheet.Cells[1, 4] = "Salary";

        //    //Format A1:D1 as bold, vertical alignment = center.
        //    oSheet.get_Range("A1", "D1").Font.Bold = true;
        //    oSheet.get_Range("A1", "D1").VerticalAlignment =
        //        Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

        //    // Create an array to multiple values at once.
        //    string[,] saNames = new string[5, 2];

        //    saNames[0, 0] = "John";
        //    saNames[0, 1] = "Smith";
        //    saNames[1, 0] = "Tom";

        //    saNames[4, 1] = "Johnson";

        //    //Fill A2:B6 with an array of values (First and Last Names).
        //    oSheet.get_Range("A2", "B6").Value2 = saNames;

        //    //Fill C2:C6 with a relative formula (=A2 & " " & B2).
        //    oRng = oSheet.get_Range("C2", "C6");
        //    oRng.Formula = "=A2 & \" \" & B2";

        //    //Fill D2:D6 with a formula(=RAND()*100000) and apply format.
        //    oRng = oSheet.get_Range("D2", "D6");
        //    oRng.Formula = "=RAND()*100000";
        //    oRng.NumberFormat = "$0.00";

        //    //AutoFit columns A:D.
        //    oRng = oSheet.get_Range("A1", "D1");
        //    oRng.EntireColumn.AutoFit();

        //    oXL.Visible = false;
        //    oXL.UserControl = false;
        //    oWB.SaveAs("c:\\test\\test505.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
        //        false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
        //        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

        //    oWB.Close();
        //    oXL.Quit();

        //    //...
        //}
    }
}
