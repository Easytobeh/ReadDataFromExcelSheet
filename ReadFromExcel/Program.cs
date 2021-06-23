using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.Text;
using System.Linq;
using System.Data;
using Newtonsoft.Json;
using DocumentFormat.OpenXml;
using System.Reflection;

namespace ReadFromExcel
{
    public class DbTable
    {
        public class PaymentAdvice
        {
            public string ID { get; set; }
            public string Name { get; set; }
            public string Salary { get; set; }
            public string Date { get; set; }
            public string Staff_Id { get; set; }
            public string Staff_Name { get; set; }
            public string Staff_Role { get; set; }
            public string Staff_Grade { get; set; }
            public string Hire_Date { get; set; }
            public string Days_Worked { get; set; }
            public string Monthly_Gross_Income { get; set; }
            public string Basic_Salary { get; set; }
            public string Housing { get; set; }
            public string Transport { get; set; }
            public string Lunch { get; set; }
            public string Utility { get; set; }
            public string Entertainment { get; set; }
            public string Education { get; set; }
            public string Dressing { get; set; }
            public string Leave_Allowance { get; set; }
            public string Car_Monetization { get; set; }
            public string NHF { get; set; }
            public string Monthly_Total_Deductions { get; set; }
            public string Vol_Cont_Pension { get; set; }
            public string Monthly_Net_Sal { get; set; }
            public string Upload_Date { get; set; }
            public string Month { get; set; }
            public string Year { get; set; }
            public string Email { get; set; }
        }
        public class PayslipStatus
        {
            public string StaffId { get; set; }
            public string Month { get; set; }
            public string Year { get; set; }
            public string Status { get; set; }
            public string Filename { get; set; }
            public string DaysGenerated { get; set; }

        }
       
    }
    class Program
    {

        static void Main(string[] args)
        {

            // ReadExcelFile();
            ReadExcelFilereal();
            //WriteExcelFile();
        }
        static void ReadExcelFilereal()
        {
            try
            {
                string strDoc = @"C:\Users\bashir.adeyemi\source\repos\ExcelToDb\ExcelToDb\Files\Payslip_Test_Table.xlsx";

                //Lets open the existing excel file and read through its content . Open the excel using openxml sdk
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(strDoc, false))
                {
                    //create the object for workbook part  
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    Sheets thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();
                    string excelResult = "";
                    int count = 0;
                    int count2 = 0;
                    List<DbTable.PaymentAdvice> Datalists = new List<DbTable.PaymentAdvice>();
                    var Staff = new DbTable.PaymentAdvice();
                    var Payslip = new DbTable.PayslipStatus();


                    //using for each loop to get the sheet from the sheetcollection  
                    foreach (Sheet thesheet in thesheetcollection)
                    {
                        //statement to get the worksheet object by using the sheet id  
                        Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;

                        SheetData thesheetdata = (SheetData)theWorksheet.GetFirstChild<SheetData>();
                        try
                        {
                            foreach (Row thecurrentrow in thesheetdata)
                            {
                                if (count != 0)
                                {
                                    foreach (Cell thecurrentcell in thecurrentrow)
                                    {
                                        count2++;
                                        //statement to take the integer value  
                                        string currentcellvalue = string.Empty;
                                        if (thecurrentcell.DataType != null)
                                        {
                                            if (thecurrentcell.DataType == CellValues.SharedString)
                                            {
                                                int id;
                                                if (Int32.TryParse(thecurrentcell.InnerText, out id))
                                                {
                                                    SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                                    if (item.Text != null)
                                                    {
                                                        //code to take the string value  
                                                        excelResult = item.Text.Text;
                                                    }
                                                    else if (item.InnerText != null)
                                                    {
                                                        currentcellvalue = item.InnerText;
                                                    }
                                                    else if (item.InnerXml != null)
                                                    {
                                                        currentcellvalue = item.InnerXml;
                                                    }
                                                }
                                            }
                                        }

                                        else
                                        {
                                            excelResult = thecurrentcell.InnerText;
                                        }

                                        switch (count2)
                                        {
                                            case 1:
                                                Staff.Staff_Id = excelResult;
                                                break;
                                            case 2:
                                                Staff.Staff_Name = excelResult;
                                                break;
                                            case 3:
                                                Staff.Staff_Role = excelResult;
                                                break;
                                            case 4:
                                                Staff.Staff_Grade = excelResult;
                                                break;
                                            case 5:
                                                Staff.Hire_Date = excelResult;
                                                break;
                                            case 6:
                                                Staff.Days_Worked = excelResult;
                                                break;
                                            case 7:
                                                Staff.Monthly_Gross_Income = excelResult;
                                                break;
                                            case 8:
                                                Staff.Basic_Salary = excelResult;
                                                break;
                                            case 9:
                                                Staff.Housing = excelResult;
                                                break;
                                            case 10:
                                                Staff.Transport = excelResult;
                                                break;
                                            case 11:
                                                Staff.Lunch = excelResult;
                                                break;
                                            case 12:
                                                Staff.Utility = excelResult;
                                                break;
                                            case 13:
                                                Staff.Entertainment = excelResult;
                                                break;
                                            case 14:
                                                Staff.Education = excelResult;
                                                break;
                                            case 15:
                                                Staff.Dressing = excelResult;
                                                break;
                                            case 16:
                                                Staff.Leave_Allowance = excelResult;
                                                break;
                                            case 17:
                                                Staff.Car_Monetization = excelResult;
                                                break;
                                            case 18:
                                                Staff.NHF = excelResult;
                                                break;
                                            case 19:
                                                Staff.Monthly_Total_Deductions = excelResult;
                                                break;
                                            case 20:
                                                Staff.Vol_Cont_Pension = excelResult;
                                                break;
                                            case 21:
                                                Staff.Monthly_Net_Sal = excelResult;
                                                break;
                                            case 22:
                                                Staff.Upload_Date = excelResult;
                                                break;
                                            case 23:
                                                Staff.Month = excelResult;
                                                break;
                                            case 24:
                                                Staff.Year = excelResult;
                                                break;
                                            case 25:
                                                {
                                                    Staff.Email = excelResult;
                                                    //Display result at the end of each row
                                                    DisplayPayslipDetails(Staff);
                                                    break;
                                                }
                                               
                                            default:
                                                break;

                                        }
                                        //switch (count2)
                                        //{
                                        //    case 1:
                                        //        Payslip.StaffId = excelResult;
                                        //        break;
                                        //    case 2:
                                        //        Payslip.Month = excelResult;
                                        //        break;
                                        //    case 3:
                                        //        Payslip.Year = excelResult;
                                        //        break;
                                        //    case 4:
                                        //        Payslip.Status = excelResult;
                                        //        break;
                                        //    case 5:
                                        //        Payslip.Filename = excelResult;
                                        //        break;
                                        //    case 6:
                                        //        Payslip.DaysGenerated = excelResult;
                                        //        break;                                            
                                        //    case 7:
                                        //        {
                                        //            Staff.Email = excelResult;
                                        //            //Display result at the end of each row
                                        //            DisplayPayslipStatus(Payslip);
                                        //            break;
                                        //        }

                                        //    default:
                                        //        break;

                                        //}
                                    }
            
                                }
                                count++;
                                count2 = 0;

                            }
                          
                            Console.ReadLine();
                        }
                        catch (Exception ex)
                        {

                            Console.WriteLine(ex.Message + "Stacktrace:" + ex.StackTrace); ;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:" + ex.Message + "," + ex.StackTrace);
            }
        }
        static void ReadExcelFile()
        {
            try
            {
                string strDoc = @"C:\Users\bashir.adeyemi\source\repos\ExcelToDb\ExcelToDb\Files\test_table.xlsx";

                //Lets open the existing excel file and read through its content . Open the excel using openxml sdk
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(strDoc, false))
                {
                    //create the object for workbook part  
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    Sheets thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();
                    StringBuilder excelResult = new StringBuilder();

                    //using for each loop to get the sheet from the sheetcollection  
                    foreach (Sheet thesheet in thesheetcollection)
                    {
                        excelResult.AppendLine("Excel Sheet Name : " + thesheet.Name);
                        excelResult.AppendLine("----------------------------------------------- ");
                        //statement to get the worksheet object by using the sheet id  
                        Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;

                        SheetData thesheetdata = (SheetData)theWorksheet.GetFirstChild<SheetData>();
                        try
                        {
                            foreach (Row thecurrentrow in thesheetdata)
                            {
                                foreach (Cell thecurrentcell in thecurrentrow)
                                {
                                    //statement to take the integer value  
                                    string currentcellvalue = string.Empty;
                                    int id;
                                    Int32.TryParse(thecurrentcell.InnerText, out id);


                                    SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                    if (item.Text != null)
                                    {
                                        //code to take the string value  
                                        excelResult.Append(item.Text.Text + " ");
                                    }
                                    else if (item.InnerText != null)
                                    {
                                        currentcellvalue = item.InnerText;
                                    }
                                    else if (item.InnerXml != null)
                                    {
                                        currentcellvalue = item.InnerXml;
                                    }




                                    excelResult.AppendLine();
                                }

                                excelResult.Append("");
                                Console.WriteLine(excelResult.ToString());
                                Console.ReadLine();
                            }
                        }
                        catch (Exception ex)
                        {

                            Console.WriteLine(ex.Message + "Stacktrace:" + ex.StackTrace); ;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:" + ex.Message + "," + ex.StackTrace);
            }
        }

        static public void DisplayPayslipStatus(DbTable.PayslipStatus payslip)
        {
            PropertyInfo[] properties = typeof(DbTable.PayslipStatus).GetProperties();
            foreach (PropertyInfo property in properties)
            {
                Console.WriteLine(property.Name + ":" + property.GetValue(payslip));
            }
            Console.WriteLine("-------------------------");
        }

        static public void DisplayPayslipDetails(DbTable.PaymentAdvice staff)
        {
            PropertyInfo[] properties = typeof(DbTable.PaymentAdvice).GetProperties();
            foreach (PropertyInfo property in properties)
            {
                Console.WriteLine(property.Name + ":" + property.GetValue(staff));
            }
            Console.WriteLine("-------------------------");
           
        }
        static void ReadExcelFilerealest()
        {
            try
            {
                string strDoc = @"C:\Users\bashir.adeyemi\source\repos\ExcelToDb\ExcelToDb\Files\test_table.xlsx";

                //Lets open the existing excel file and read through its content . Open the excel using openxml sdk
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(strDoc, false))
                {
                    //create the object for workbook part  
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    Sheets thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();
                    StringBuilder excelResult = new StringBuilder();

                    //using for each loop to get the sheet from the sheetcollection  
                    foreach (Sheet thesheet in thesheetcollection)
                    {
                        excelResult.AppendLine("Excel Sheet Name : " + thesheet.Name);
                        excelResult.AppendLine("----------------------------------------------- ");
                        //statement to get the worksheet object by using the sheet id  
                        Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;

                        SheetData thesheetdata = (SheetData)theWorksheet.GetFirstChild<SheetData>();
                        try
                        {
                            foreach (Row thecurrentrow in thesheetdata)
                            {
                                foreach (Cell thecurrentcell in thecurrentrow)
                                {
                                    //statement to take the integer value  
                                    string currentcellvalue = string.Empty;
                                    if (thecurrentcell.DataType != null)
                                    {
                                        if (thecurrentcell.DataType == CellValues.SharedString)
                                        {
                                            int id;
                                            if (Int32.TryParse(thecurrentcell.InnerText, out id))
                                            {
                                                SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                                if (item.Text != null)
                                                {
                                                    //code to take the string value  
                                                    excelResult.Append(item.Text.Text + " ");
                                                }
                                                else if (item.InnerText != null)
                                                {
                                                    currentcellvalue = item.InnerText;
                                                }
                                                else if (item.InnerXml != null)
                                                {
                                                    currentcellvalue = item.InnerXml;
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        excelResult.Append(Convert.ToString(thecurrentcell.InnerText) + " ");
                                    }
                                }
                                excelResult.AppendLine();
                            }

                            excelResult.Append("");
                            Console.WriteLine(excelResult.ToString());
                            Console.ReadLine();
                        }
                        catch (Exception ex)
                        {

                            Console.WriteLine(ex.Message + "Stacktrace:" + ex.StackTrace); ;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:" + ex.Message + "," + ex.StackTrace);
            }
        }

        static void WriteExcelFile()
        {
            List<UserDetails> persons = new List<UserDetails>()
           {
               new UserDetails() {ID="1001", Name="ABCD", City ="City1", Country="USA"},
               new UserDetails() {ID="1002", Name="PQRS", City ="City2", Country="INDIA"},
               new UserDetails() {ID="1003", Name="XYZZ", City ="City3", Country="CHINA"},
               new UserDetails() {ID="1004", Name="LMNO", City ="City4", Country="UK"},
          };

            string strDoc = @"C:\Users\bashir.adeyemi\source\repos\ExcelToDb\ExcelToDb\Files\new_table.xlsx";

            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(persons), (typeof(DataTable)));

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(strDoc, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };

                sheets.Append(sheet);

                Row headerRow = new Row();

                List<String> columns = new List<string>();
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
                        cell.CellValue = new CellValue(dsrow[col].ToString());
                        newRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(newRow);
                }

                workbookPart.Workbook.Save();
            }
        }
    }

    internal class UserDetails
    {
        public UserDetails()
        {
        }
       
        public string ID { get; set; }
        public string Name { get; set; }
        public string City { get; set; }
        public string Country { get; set; }
    }

}