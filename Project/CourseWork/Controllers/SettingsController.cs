using System.Web.Mvc;
using CourseWork.Models;

namespace CourseWork.Controllers
{
    public class SettingsController : Controller
    {
        // GET: Settings
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(string[] chk)
        {
            if (Request["chk1"] == "on") SettingsModels.Abs = true; else SettingsModels.Abs = false;
            if (Request["chk2"] == "on") SettingsModels.RawScore = true; else SettingsModels.RawScore = false;
            if (Request["chk3"] == "on") SettingsModels.Perc = true; else SettingsModels.Perc = false;
            if (Request["chk4"] == "on") SettingsModels.Z = true; else SettingsModels.Z = false;
            if (Request["chk5"] == "on") SettingsModels.T = true; else SettingsModels.T = false;
            if (Request["chk6"] == "on") SettingsModels.Ckvar25 = true; else SettingsModels.Ckvar25 = false;
            if (Request["chk7"] == "on") SettingsModels.Ckvar75 = true; else SettingsModels.Ckvar75 = false;
            if (Request["chk8"] == "on") SettingsModels.Table = true; else SettingsModels.Table = false;
            if (Request["radio"] == "1") SettingsModels.Sort = true; else SettingsModels.Sort = false;

            return RedirectToAction("Index");
        }
    }
}