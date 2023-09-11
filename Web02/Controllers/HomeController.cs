using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Web02.Utilities;

namespace Web02.Controllers
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

        public ActionResult GenerarActaView()
        {
            var utilities = new Utilities.Utilities();
            MemoryStream stream = utilities.GenerarActa();

            Response.ContentType = "application/msword";
            Response.AddHeader("content-disposition", $"attachment; filename=Acta_{DateTime.Now:ddMMyyyy_HHmm}.docx");

            stream.Seek(0, SeekOrigin.Begin);
            stream.CopyTo(Response.OutputStream);

            return new EmptyResult();
        }

    }
}