using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using WagnerCable.Models;
using WagnerCable.Utilities;

namespace WagnerCable.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult ImportExcel()
        {
            return View();
        }
        [HttpPost]
        public ActionResult ImportExcel(HttpPostedFileBase PostedFile)
        {
            try
            {

                if (PostedFile.ContentLength > 0)
                {
                    string extension = System.IO.Path.GetExtension(PostedFile.FileName).ToLower();
                    string query = null;
                    string connString = "";
                    string fileName = Guid.NewGuid().ToString();
                    List<ExcelViewModel> list = new List<ExcelViewModel>();
                    List<ExcelViewModel> newList = new List<ExcelViewModel>();


                    string[] validFileTypes = { ".xls", ".xlsx", ".csv" };

                    string path1 = string.Format("{0}/{1}", Server.MapPath("~/Content/Uploads"), fileName);
                    if (!Directory.Exists(path1))
                    {
                        Directory.CreateDirectory(Server.MapPath("~/Content/Uploads"));
                    }
                    if (validFileTypes.Contains(extension))
                    {
                        if (System.IO.File.Exists(path1))
                        {
                            System.IO.File.Delete(path1);
                        }
                        PostedFile.SaveAs(path1);
                        string data = "";
                        //Create COM Objects. Create a COM object for everything that is referenced
                        Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path1);
                        Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                        Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

                        int rowCount = xlRange.Rows.Count;
                        int colCount = xlRange.Columns.Count;

                        //iterate over the rows and columns and print to the console as it appears in the file
                        //excel is not zero based!!
                        for (int i = 2; i <= rowCount; i++)
                        {
                            ExcelViewModel model = new ExcelViewModel();
                            //either collect data cell by cell or DO you job like insert to DB 
                            if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                            {
                                model.Email = xlRange.Cells[i, 1].Value2.ToString();
                                SendEmail(model.Email);
                            }
                           
                        }
                       
                    }
                    else
                    {
                        ViewBag.Error = "Please Upload Files in .xls, .xlsx or .csv format";

                    }

                }

                return View();
            }
            catch (Exception ex)
            {
                ViewBag.Message = "Failed to Upload File";
                return View();
            }
        }


        private static void SendEmail(string emails, string mailKonu = "Mutlu yıllar", string content = "lumhar.com/2022")
        {
            StringBuilder mail = new StringBuilder();

            Random rnd = new Random();


            UtilityManager.SendEmail(emails, mailKonu, content);

        }

    }
}