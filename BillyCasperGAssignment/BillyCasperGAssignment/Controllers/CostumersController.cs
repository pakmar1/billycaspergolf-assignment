using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BillyCasperGAssignment.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MySql.Data.MySqlClient;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Excel = Microsoft.Office.Interop.Excel;

namespace BillyCasperGAssignment.Controllers
{
    public class CostumersController : Controller
    {

        private IHostingEnvironment _hostingEnvironment;
        public CostumersController(IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;

        }

        // GET: /<controller>/
        public IActionResult Index()
        {
            return View();
        }


        public IActionResult Result()
        {
            BillyCasperContext context = HttpContext.RequestServices.GetService(typeof(BillyCasperContext)) as BillyCasperContext;
            return View(context.GetAllCostumers());
        }

        //[HttpPost]
        public ActionResult Import()
        {
            BillyCasperContext context = HttpContext.RequestServices.GetService(typeof(BillyCasperContext)) as BillyCasperContext;

            IFormFile file = Request.Form.Files[0];
            string folderName = "Upload";
            string webRootPath = _hostingEnvironment.WebRootPath;
            string newPath = Path.Combine(webRootPath, folderName);
            StringBuilder sb = new StringBuilder();
            if (!Directory.Exists(newPath))
            {
                Directory.CreateDirectory(newPath);
            }
            if (file.Length > 0)
            {

                string sFileExtension = Path.GetExtension(file.FileName).ToLower();
                ISheet sheet;
                string fullPath = Path.Combine(newPath, file.FileName);
                using (var stream = new FileStream(fullPath, FileMode.Create))
                {
                    file.CopyTo(stream);
                    stream.Position = 0;
                    if (sFileExtension == ".xls")
                    {
                        HSSFWorkbook hssfwb = new HSSFWorkbook(stream); //This will read the Excel 97-2000 formats  
                        sheet = hssfwb.GetSheetAt(1); //get first sheet from workbook  
                    }
                    else
                    {
                        XSSFWorkbook hssfwb = new XSSFWorkbook(stream); //This will read 2007 Excel format  
                        sheet = hssfwb.GetSheetAt(1); //get first sheet from workbook   
                    }
                    IRow headerRow = sheet.GetRow(0); //Get Header Row

                    //context.AddCost();
                    context.AddCostumers(sheet);

                }
            }
            return this.Content(sb.ToString());
        }

        public async Task<ActionResult> Export()
        {
            string sWebRootFolder = _hostingEnvironment.WebRootPath;
            string sFileName = @"result.xlsx";
            string URL = string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, sFileName);
            FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
            var memory = new MemoryStream();
            using (var fs = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook;
                workbook = new XSSFWorkbook();
                ISheet excelSheet = workbook.CreateSheet("Demo");
                List<Costumer> list = new List<Costumer>();
                IRow row;int num = 1;
                using (MySqlConnection conn = createConnect())
                {
                    conn.Open();
                    MySqlCommand cmd = new MySqlCommand("SELECT * FROM Costumer order by ID", conn);
                    using (MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            list.Add(new Costumer()
                            {
                                ID = reader.GetInt32("ID"),
                                CreatedOn = reader.GetDateTime("CreatedOn"),
                                ModifiedOn = reader.GetDateTime("ModifiedOn"),
                                Costumer_LastName = reader.GetString("Costumer_LastName"),
                                Costumer_FirstName = reader.GetString("Costumer_FirstName"),
                                AddressLine1 = reader.GetString("Costumer_AddressLine1"),
                                Costumer_City = reader.GetString("Costumer_City"),
                                Costumer_State = reader.GetString("Costumer_State"),
                                Costumer_zip = reader.GetString("Costumer_Zip"),
                                Costumer_Homephone = reader.GetString("Costumer_HomePhone"),
                                Costumer_InternetEmail = reader.GetString("Costumer_InternetEmail")
                            });
                        }
                    }

                    row = excelSheet.CreateRow(num);
                    row.CreateCell(0).SetCellValue("ID");
                    row.CreateCell(1).SetCellValue("CreatedOn");
                    row.CreateCell(2).SetCellValue("ModifiedOn");
                    row.CreateCell(3).SetCellValue("Costumer_LastName");
                    row.CreateCell(4).SetCellValue("Costumer_FirstName");
                    row.CreateCell(5).SetCellValue("AddressLine1");
                    row.CreateCell(6).SetCellValue("Costumer_City");
                    row.CreateCell(7).SetCellValue("Costumer_State");
                    row.CreateCell(8).SetCellValue("Costumer_zip");
                    row.CreateCell(9).SetCellValue("Costumer_Homephone");
                    row.CreateCell(10).SetCellValue("Costumer_InternetEmail");

                    foreach (var item in list)
                    {
                        row = excelSheet.CreateRow(num);
                        row.CreateCell(0).SetCellValue(item.ID);
                        row.CreateCell(1).SetCellValue(item.CreatedOn);
                        row.CreateCell(2).SetCellValue(item.ModifiedOn);
                        row.CreateCell(3).SetCellValue(item.Costumer_LastName);
                        row.CreateCell(4).SetCellValue(item.Costumer_FirstName);
                        row.CreateCell(5).SetCellValue(item.AddressLine1);
                        row.CreateCell(6).SetCellValue(item.Costumer_City);
                        row.CreateCell(7).SetCellValue(item.Costumer_State);
                        row.CreateCell(8).SetCellValue(item.Costumer_zip);
                        row.CreateCell(9).SetCellValue(item.Costumer_Homephone);
                        row.CreateCell(10).SetCellValue(item.Costumer_InternetEmail);
                        num++;
                    }


                }
            }
            using (var stream = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Open))
            {
                await stream.CopyToAsync(memory);
            }
            memory.Position = 0;
            return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", sFileName);
        }


    }
}