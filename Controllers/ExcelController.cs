using adodotnet.DAL;
using adodotnet.Utility;
using ClosedXML.Excel;
using ExcelFileImportExport.Models;
using Microsoft.AspNetCore.Mvc;
//using MySql.Data.MySqlClient;
using MySqlConnector;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Text;

namespace ExcelFileImportExport.Controllers
{
    public class ExcelController : Controller
    {

        private readonly string connectionString;
        private CustomerDAL customerDAL = null;

        public ExcelController()
        {
            connectionString = ConnectionString._ConnectionString;
            this.customerDAL = new CustomerDAL();
        }
        //public IActionResult Index()
        //{
        //    IEnumerable<ExcelCustomer> ExcelCustomerone = customerDAL.GetAllExcelCustomers();
        //    return View(ExcelCustomerone);
        //}

        public IActionResult Index(int PageNumber)
        {

            var customers = customerDAL.GetAllExcelCustomers();
            ViewBag.TotalPages = Math.Ceiling(customers.Count() / 10.0);
            ViewBag.PageNumber = PageNumber;
            customers = customers.Skip((PageNumber - 1) * 10).Take(10).ToList();

            return View(customers);
        }

        public IActionResult ImportExcelFile()
        {
            return View();
        }

        [HttpPost]
        public IActionResult ImportExcelFile(IFormFile formFile)
        {

            try
            {
                
                    var mainpath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "UploadExcelFile");
                    if (!Directory.Exists(mainpath))
                    {
                        Directory.CreateDirectory(mainpath);
                    }
                    var filePath = Path.Combine(mainpath, formFile.FileName);
                    using (FileStream stream = new FileStream(filePath, FileMode.Create))
                    {
                        formFile.CopyTo(stream);
                    }
                    var fileName = formFile.FileName;
                    string extension = Path.GetExtension(fileName);
                    string conString = string.Empty;
                    switch (extension)
                    {
                        case ".xls":
                            conString = "Provider=Microsoft.ACE.OLEDB.12.0;; Data Source=" + filePath + ";Extended Properties='Excel 8.0; HDR=Yes'";
                            break;
                        case ".xlsx":
                            conString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + filePath + ";Extended Properties='Excel 8.0; HDR=Yes'";
                            break;
                        case ".csv":
                            conString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + Path.GetDirectoryName(filePath) + ";Extended Properties='Text; HDR=Yes; FMT=CSVDelimited'";
                            break;
                }
                    DataTable dt = new DataTable();
                    conString = string.Format(conString, filePath);
                    using (OleDbConnection conExcel = new OleDbConnection(conString))
                    {
                        using (OleDbCommand cmdExcel = conExcel.CreateCommand())
                        {
                            using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                            {
                                cmdExcel.Connection = conExcel;
                                conExcel.Open();
                                DataTable dtExcelSchema = conExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                                string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                                cmdExcel.CommandText = "SELECT * FROM [" + sheetName + "]";
                                odaExcel.SelectCommand = cmdExcel;
                                odaExcel.Fill(dt);
                                conExcel.Close();
                            }
                        }
                    }

                    conString = connectionString;
                    using (MySqlConnection con = new MySqlConnection(connectionString))
                    {
                        con.Open();
                        foreach (DataRow row in dt.Rows)
                        {
                            MySqlCommand cmd = new MySqlCommand("INSERT INTO exceldata (FirstName, LastName, Gender, Country, Age) VALUES (@FirstName, @LastName, @Gender, @Country, @Age);", con);
                            cmd.Parameters.AddWithValue("@FirstName", row["FirstName"]);
                            cmd.Parameters.AddWithValue("@LastName", row["LastName"]);
                            cmd.Parameters.AddWithValue("@Gender", row["Gender"]);
                            cmd.Parameters.AddWithValue("@Country", row["Country"]);
                            cmd.Parameters.AddWithValue("@Age", row["Age"]);

                            cmd.ExecuteNonQuery();
                        }
                        con.Close();
                    }
                    TempData["Message"] = "File Imported Successfully";
                    return RedirectToAction("Index");
  
            }
            catch (Exception ex) 
            { 
                string message = ex.Message;
            }
            return View();
        }

        public IActionResult ExportExcelFile()
        {
            IEnumerable<ExcelCustomer> customers = customerDAL.GetAllExcelCustomers();

            try
            {
                using(var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("customers");
                    var currentRow = 1;
                    worksheet.Cell(currentRow, 1).Value = "Id";
                    worksheet.Cell(currentRow, 2).Value = "FirstName";
                    worksheet.Cell(currentRow, 3).Value = "LastName";
                    worksheet.Cell(currentRow, 4).Value = "Gender";
                    worksheet.Cell(currentRow, 5).Value = "Country";
                    worksheet.Cell(currentRow, 6).Value = "Age";

                    foreach (var cst in customers)
                    {
                          currentRow++;
                          worksheet.Cell(currentRow, 1).Value = cst.Id;
                          worksheet.Cell(currentRow,2).Value = cst.FirstName;
                          worksheet.Cell(currentRow, 3).Value = cst.LastName;
                          worksheet.Cell(currentRow, 4).Value = cst.Gender;
                          worksheet.Cell(currentRow, 5).Value = cst.Country;
                          worksheet.Cell(currentRow, 6).Value = cst.Age;                        
                    }

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        var content= stream.ToArray();
                        return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CustomerInfo.xlsx");
                    }
                }
            }
            catch (Exception ex)
            {
                string message = ex.Message;
            }
            return View();
        }

        public IActionResult CSV()
        {
            IEnumerable<ExcelCustomer> customers = customerDAL.GetAllExcelCustomers();


            var builder = new StringBuilder();
            builder.AppendLine("Id,FirstName,LastName,Gender,Country,Age");
            foreach (var customer in customers)
            {
                builder.AppendLine($"{customer.Id},{customer.FirstName},{customer.LastName},{customer.Gender},{customer.Country},{customer.Age}");
            }
            return File(Encoding.UTF8.GetBytes(builder.ToString()),"text/csv","CustomerinfoCSV.csv");
        }
    }
}
