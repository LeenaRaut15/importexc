using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using OfficeOpenXml;
using importexc.Data;
using importexc.Models;
using System.IO;

namespace importexc.Controllers.Excel
{
    public class ExcelController : Controller
    {
        // GET: Excel
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult Upload(FormCollection formCollection)
        {
            var Excel = new List<tbl_Excel>();
            if (Request != null)
            {
                HttpPostedFileBase file = Request.Files["UploadFile"];
                if ((file != null) && (file.ContentLength > 0) && !string.IsNullOrEmpty(file.FileName))
                {
                    string fileName = file.FileName;
                    string fileContentType = file.ContentType;
                    byte[] fileBytes = new byte[file.ContentLength];
                    var data = file.InputStream.Read(fileBytes, 0, Convert.ToInt32(file.ContentLength));
                    using (var package = new ExcelPackage(file.InputStream))
                    {
                        var currentSheet = package.Workbook.Worksheets;
                        var workSheet = currentSheet.First();
                        var noOfCal = workSheet.Dimension.End.Column;
                        var noOfRow = workSheet.Dimension.End.Row;

                        using (EXCELEntities db = new EXCELEntities())
                        {
                            for (int rowIterator = 2; rowIterator <= noOfRow; rowIterator++)
                            {
                                var user = new tbl_Excel();
                                user.HospitalName = workSheet.Cells[rowIterator, 1].Value.ToString();
                                user.AddmissionFee = workSheet.Cells[rowIterator, 2].Value.ToString();
                                user.Address = workSheet.Cells[rowIterator, 3].Value.ToString();
                                db.tbl_Excel.Add(user);
                                db.SaveChanges();
                            }
                        }

                    }
                }
            }

            return View("Index");
        }
    }
}
    
