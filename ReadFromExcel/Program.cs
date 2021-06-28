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
using System.IO;

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
            public string Nuban { get; set; }
            public string Days_Worked { get; set; }
            public string Annual_Gross_Income { get; set; }
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
            public string Monthly_Pension { get; set; }
            public string PAYE { get; set; }
            public string Monthly_Net_Sal { get; set; }
            public string Upload_Date { get; set; }
            public string Month { get; set; }
            public string Year { get; set; }
            public string Email { get; set; }
            public string Total_Deduction { get; internal set; }
            public string Leave_Allowance_Per_Month { get; internal set; }
            public string Gross_Total { get; internal set; }
            public string Total_Per_Month { get; internal set; }
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
                string strDoc = @"C:\Users\bashir.adeyemi\source\repos\ReadDataFromExcelSheet_ConsoleApp\ReadFromExcel\Payslip_Test_Table.xlsx";
                
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

                    //Output pdf file converted
          
                    //Create path


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
                                                Staff.Nuban = excelResult;
                                                break;
                                            case 7:
                                                Staff.Days_Worked = excelResult;
                                                break;
                                            case 8:
                                                Staff.Annual_Gross_Income = excelResult;
                                                break;
                                            case 9:
                                                Staff.Monthly_Gross_Income = excelResult;
                                                break;
                                            case 10:
                                                Staff.Basic_Salary = excelResult;
                                                break;
                                            case 11:
                                                Staff.Housing = excelResult;
                                                break;
                                            case 12:
                                                Staff.Transport = excelResult;
                                                break;
                                            case 13:
                                                Staff.Lunch = excelResult;
                                                break;
                                            case 14:
                                                Staff.Utility = excelResult;
                                                break;
                                            case 15:
                                                Staff.Entertainment = excelResult;
                                                break;
                                            case 16:
                                                Staff.Education = excelResult;
                                                break;
                                            case 17:
                                                Staff.Dressing = excelResult;
                                                break;
                                            case 18:
                                                Staff.Leave_Allowance = excelResult;
                                                break;
                                            case 19:
                                                Staff.Car_Monetization = excelResult;
                                                break;
                                            case 20:
                                                Staff.Total_Per_Month = excelResult;
                                                break;
                                            case 21:
                                                Staff.Gross_Total = excelResult;
                                                break;
                                            case 22:
                                                Staff.PAYE = excelResult;
                                                break;
                                            case 23:
                                                Staff.Monthly_Pension = excelResult;
                                                break;
                                            case 24:
                                                Staff.Leave_Allowance_Per_Month = excelResult;
                                                break;
                                            case 25:
                                                Staff.NHF = excelResult;
                                                break;
                                            case 26:
                                                Staff.Vol_Cont_Pension = excelResult;
                                                break;
                                            case 27:
                                                Staff.Total_Deduction = excelResult;
                                                break;
                                            case 28:
                                                Staff.Monthly_Net_Sal = excelResult;
                                                break;
                                            case 29:
                                                Staff.Month = excelResult;
                                                break;
                                            case 30:
                                                Staff.Year = excelResult;
                                                break;
                                            case 31:
                                                {
                                                    Staff.Email = excelResult;
                                                    //Display result at the end of each row
                                                   // DisplayPayslipDetails(Staff);
                                                    EditHtmlFile(Staff);
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
       
        static string CreateFileName(string staffName)
        {
            return $"{staffName}_{DateTime.Now}";
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
       static void EditHtmlFile(DbTable.PaymentAdvice uploadedPayslip)
        {
            string fileName = "";
            string PDFPath = "C:\\Users\\bashir.adeyemi\\source\\repos\\ReadDataFromExcelSheet_ConsoleApp\\Output";
            string htmlTemplate = "C:\\Users\\bashir.adeyemi\\source\\repos\\ReadDataFromExcelSheet_ConsoleApp\\PDFTemplate.html";

            try
            {
                string html = File.ReadAllText(htmlTemplate)
                    .Replace("{Staff_Name}", uploadedPayslip.Staff_Name)
                    .Replace("{Staff_ID}", uploadedPayslip.Staff_Id)
                    .Replace("{Staff_Grade}", uploadedPayslip.Staff_Grade)
                    .Replace("{Staff_Role}", uploadedPayslip.Staff_Role)
                    .Replace("{Month}", uploadedPayslip.Month)
                .Replace("{Days_Worked}", uploadedPayslip.Days_Worked)
                .Replace("{Basic_Salary}", uploadedPayslip.Basic_Salary)
                .Replace("{Housing_Allowance}", uploadedPayslip.Housing)
                .Replace("{Transport_Allowance}", uploadedPayslip.Transport)
                .Replace("{Utility}", uploadedPayslip.Utility)
                .Replace("{Entertainment_Allowance}", uploadedPayslip.Entertainment)
                .Replace("{Lunch}", uploadedPayslip.Lunch)
                .Replace("{Education}", uploadedPayslip.Education)
                .Replace("{Dressing_Allowance}", uploadedPayslip.Dressing)
                .Replace("{Leave_Allowance}", uploadedPayslip.Leave_Allowance)
                .Replace("{Car_Monetization}", uploadedPayslip.Car_Monetization)
                .Replace("{Earnings}", uploadedPayslip.Monthly_Gross_Income)
                .Replace("{PAYE_Tax}", uploadedPayslip.PAYE)
                .Replace("{Pension}", uploadedPayslip.Monthly_Pension)
                .Replace("{Leave_Allowance_PM}", uploadedPayslip.Leave_Allowance_Per_Month)
                .Replace("{NHF}", uploadedPayslip.NHF)
                .Replace("{Voluntary_Contribution}", uploadedPayslip.Vol_Cont_Pension)
                .Replace("{Total_Deductions}", uploadedPayslip.Total_Deduction)
                .Replace("{Net_Monthly_Pay}", uploadedPayslip.Monthly_Net_Sal);


                //call method that generate unique name for each file
                fileName = uploadedPayslip.Staff_Name + ".html";//CreateFileName(uploadedPayslip.Name);

                //save HtmFile for each staff
               File.WriteAllText(PDFPath + "\\" + fileName, html);

            }
            catch (Exception e)
            {
                var ex = e;
            }
          //  return PDFPath + "\\" + fileName;

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