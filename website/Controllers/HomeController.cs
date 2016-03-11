using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Whom;

namespace website.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            var derp = Program.IndexViewModel();
            return View(derp);
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
    }
}