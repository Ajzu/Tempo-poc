using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Text;

namespace well_estimation_tool_grid_poc.Helpers
{
    public class ExcelOpenXMLHelper
    {
        public void TestOpenXMLExcel()
        {
            try
            {
                string fileName = @"C:\Users\arsing5\Desktop\Test\well-estimation-tool-grid-poc\well-estimation-tool-grid-poc\test\HelloWorld.xlsx";// "Your Excel file physical path";
                //string fileName = @"c:\path\to\my\file.xlsx";


                StringBuilder sb1 = new StringBuilder();
                StringBuilder sb2 = new StringBuilder();

                using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                    {
                        if (doc != null)
                        {
                            WorkbookPart workbookPart = doc.WorkbookPart;
                            SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                            SharedStringTable sst = sstpart.SharedStringTable;

                            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                            Worksheet sheet = worksheetPart.Worksheet;

                            var cells = sheet.Descendants<Cell>();
                            var rows = sheet.Descendants<Row>();

                            Console.WriteLine("Row count = {0}", rows.LongCount());
                            Console.WriteLine("Cell count = {0}", cells.LongCount());

                            // One way: go through each cell in the sheet
                            foreach (Cell cell in cells)
                            {
                                if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                                {
                                    
                                    int ssid = int.Parse(cell.CellValue.Text);
                                    string str = sst.ChildElements[ssid].InnerText;
                                    Console.WriteLine("Shared string {0}: {1}", ssid, str);
                                    sb1.AppendLine("Shared string "+ ssid + " : "+ str);
                                }
                                else if (cell.CellValue != null)
                                {
                                    Console.WriteLine("Cell contents: {0}", cell.CellValue.Text);
                                    sb1.AppendLine("Cell contents: " + cell.CellValue.Text);
                                }
                            }

                            // Or... via each row
                            foreach (Row row in rows)
                            {
                                foreach (Cell c in row.Elements<Cell>())
                                {
                                    if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                                    {
                                        int ssid = int.Parse(c.CellValue.Text);
                                        string str = sst.ChildElements[ssid].InnerText;
                                        Console.WriteLine("Shared string {0}: {1}", ssid, str);
                                        sb2.AppendLine("Shared string " + ssid + " : " + str);
                                    }
                                    else if (c.CellValue != null)
                                    {
                                        Console.WriteLine("Cell contents: {0}", c.CellValue.Text);
                                        sb2.AppendLine("Cell contents: "+ c.CellValue.Text);
                                    }
                                }
                            }
                        }
                    }
                }

                string output1 = sb1.ToString();
                string output2 = sb2.ToString();
            }
            catch (Exception ex)
            {
                // throw;
            }
        }

        public void TestOpenXMLExcelInDataTable()
        {
            try
            {
                //i want to import excel to data table
                DataTable  dt = new DataTable();
                string path = @"C:\Users\arsing5\Desktop\Test\well-estimation-tool-grid-poc\well-estimation-tool-grid-poc\test\HelloWorld.xlsx";// "Your Excel file physical path";
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
                {
                    WorkbookPart workbookPart = document.WorkbookPart;
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                    OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                    //row counter
                    int rcnt = 0;

                    while (reader.Read())
                    {


                        //find xml row element type 
                        //to understand the element type you can change your excel file eg : test.xlsx to test.zip
                        //and inside that you may observe the elements in xl/worksheets/sheet.xml
                        //that helps to understand openxml better
                        if (reader.ElementType == typeof(Row))
                        {

                            //create data table row type to be populated by cells of this row
                            DataRow tempRow = dt.NewRow();



                            //***** HANDLE THE SECOND SENARIO*****
                            //if row has attribute means it is not a empty row
                            if (reader.HasAttributes)
                            {

                                //read the child of row element which is cells

                                //here first element
                                reader.ReadFirstChild();



                                do
                                {
                                    //find xml cell element type 
                                    if (reader.ElementType == typeof(Cell))
                                    {
                                        Cell c = (Cell)reader.LoadCurrentElement();

                                        string cellValue;


                                        int actualCellIndex = CellReferenceToIndex(c);

                                        if (c.DataType != null && c.DataType == CellValues.SharedString)
                                        {
                                            SharedStringItem ssi = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(c.CellValue.InnerText));

                                            cellValue = ssi.Text.Text;
                                        }
                                        else
                                        {
                                            cellValue = c.CellValue.InnerText;
                                        }



                                        //if row index is 0 its header so columns headers are added & also can do some headers check incase
                                        if (rcnt == 0)
                                        {
                                            dt.Columns.Add(cellValue);
                                        }
                                        else
                                        {
                                            // instead of tempRow[c.CellReference] = cellValue;
                                            tempRow[actualCellIndex] = cellValue;
                                        }



                                    }


                                }
                                while (reader.ReadNextSibling());


                                //if its not the header row so append rowdata to the datatable
                                if (rcnt != 0)
                                {
                                    dt.Rows.Add(tempRow);
                                }

                                rcnt++;


                            }


                        }





                    }


                }
            }
            catch (Exception ex)
            {
                //throw;
            }
        }

        private static int CellReferenceToIndex(Cell cell)
        {
            if(cell !=null)
            {
                int index = 0;
                //string reference = cell.CellReference.ToString().ToUpper();
                string reference = Convert.ToString(cell.CellReference).ToUpper();
                foreach (char ch in reference)
                {
                    if (Char.IsLetter(ch))
                    {
                        int value = (int)ch - (int)'A';
                        index = (index == 0) ? value : ((index + 1) * 26) + value;
                    }
                    else
                        return index;
                }
                return index;
            }
            else
            {
                return -1;
            }
            
        }



        //  below are the functions to write data in excel file

        public void CallExportDataSet()
        {
            try
            {
                DataTable EmployeeDetails = new DataTable("EmployeeDetails");
                //to create the column and schema
                DataColumn EmployeeID = new DataColumn("EmpID", typeof(Int32));
                EmployeeDetails.Columns.Add(EmployeeID);
                DataColumn EmployeeName = new DataColumn("EmpName", typeof(string));
                EmployeeDetails.Columns.Add(EmployeeName);
                DataColumn EmployeeMobile = new DataColumn("EmpMobile", typeof(string));
                EmployeeDetails.Columns.Add(EmployeeMobile);
                //to add the Data rows into the EmployeeDetails table
                EmployeeDetails.Rows.Add(1001, "Andrew", "9000322579");
                EmployeeDetails.Rows.Add(1002, "Briddan", "9081223457");


                DataTable SalaryDetails = new DataTable("SalaryDetails");
                //to create the column and schema
                DataColumn SalaryId = new DataColumn("SalaryID", typeof(Int32));
                SalaryDetails.Columns.Add(SalaryId);
                DataColumn empId = new DataColumn("EmployeeID", typeof(Int32));
                SalaryDetails.Columns.Add(empId);
                DataColumn empName = new DataColumn("EmployeeName", typeof(string));
                SalaryDetails.Columns.Add(empName);
                DataColumn SalaryPaid = new DataColumn("Salary", typeof(Int32));
                SalaryDetails.Columns.Add(SalaryPaid);
                //to add the Data rows into the SalaryDetails table
                SalaryDetails.Rows.Add(10001, 1001, "Andrew", 42000);
                SalaryDetails.Rows.Add(10002, 1002, "Briddan", 30000);

                //to create the object for DataSet
                DataSet dataSet = new DataSet();
                //Adding DataTables into DataSet
                dataSet.Tables.Add(EmployeeDetails);
                dataSet.Tables.Add(SalaryDetails);
                //By using index position, we can fetch the DataTable from DataSet, here first we added the Employee table so the index position of this table is 0, let's see the following code below
                //retrieving the DataTable from dataset using the Index position
                //foreach (DataRow row in dataSet.Tables[0].Rows)
                //                {
                //                    Console.WriteLine(row["EmpID"] + ", " + row["EmpName"] + ", " + row["EmpMobile"]);
                //                }
                //                Then second table we added was SalaryDetails table which the index position was 1, now we fetching this second table by using the name, so we fetching the DataTable from DataSet using the name of the table name "SalaryDetails",
                ////retrieving DataTable from the DataSet using name of the table
                //foreach (DataRow row in dataSet.Tables["SalaryDetails"].Rows)
                //                {
                //                    Console.WriteLine(row["SalaryID"] + ", " + row["EmployeeID"] + ", " + row["EmployeeName"] + ", " + row["Salary"]);
                //                }

                string destinationPath = @"C:\Users\arsing5\Desktop\Test\well-estimation-tool-grid-poc\well-estimation-tool-grid-poc\test\MyEmployeesList.xlsx";
                ExportDataSet(dataSet, destinationPath);
            }
            catch (Exception ex)
            {

                throw;
            }
        }
        private void ExportDataSet(DataSet ds, string destination)
        {
            try
            {
                using (var workbook = SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
                {
                    var workbookPart = workbook.AddWorkbookPart();

                    workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                    workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

                    foreach (System.Data.DataTable table in ds.Tables)
                    {

                        var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                        var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                        sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                        DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                        string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                        uint sheetId = 1;
                        if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                        {
                            sheetId =
                                sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                        }

                        DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                        sheets.Append(sheet);

                        DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                        List<String> columns = new List<string>();
                        foreach (System.Data.DataColumn column in table.Columns)
                        {
                            columns.Add(column.ColumnName);

                            DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                            cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
                            headerRow.AppendChild(cell);
                        }


                        sheetData.AppendChild(headerRow);

                        foreach (System.Data.DataRow dsrow in table.Rows)
                        {
                            DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                            foreach (String col in columns)
                            {
                                DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                                newRow.AppendChild(cell);
                            }

                            sheetData.AppendChild(newRow);
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                //throw;
            }
        }
    }
}
