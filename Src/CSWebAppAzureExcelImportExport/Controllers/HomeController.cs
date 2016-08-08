using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CSWebAppAzureExcelImportExport.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            return View();
        }

        public string ImportToDB()
        {
           return Helper.ExcelImportToDB();
        }


        public string ExportToExcel()
        {
            return Helper.DBExportToExcel();
        }
    }
}