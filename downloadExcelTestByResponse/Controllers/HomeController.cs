using downloadExcelByAsposeComparisonOfAjaxBeginFormAndHtmlBeginForm.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace downloadExcelByAsposeComparisonOfAjaxBeginFormAndHtmlBeginForm.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult DownloadExcel()
        {
           var asposeExcel= new AsposeExcel();

            var downloadExcel=asposeExcel.DownloadExcel();
     
            return File(downloadExcel, "application/vnd.ms-excel", "FileName.xlsx");
        }

    }
}