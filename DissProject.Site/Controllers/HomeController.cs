using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DissProject.Repository;
using DissProject.Models;
using WebMatrix.WebData;
using System.Web.Security;

namespace DissProject.Site.Controllers
{
    [DissProject.Site.Filters.InitializeSimpleMembership]
    public class HomeController : Controller
    {
        IUowData _db;

        public HomeController()
        {
            this._db = new UowData();
        }

        public ActionResult Index()
        {
            int remove = 2;
            string currentUser = User.Identity.Name;
            if (String.IsNullOrEmpty(currentUser))
            {
                ViewBag.IsApproved = remove;
                return View();
            }

            UserProfile person = _db.UserProfiles.All()
                               .Where(x => x.UserName == currentUser)
                               .Select(x => x).SingleOrDefault();
            remove = (person.IsApproved.Value && person.IsApproved != null) ? 1 : 2;
            if (remove == 2)
            {
                ViewBag.IsApproved = remove;
                return View();
            }

            ViewBag.IsApproved = remove;
            var roles = (SimpleRoleProvider)Roles.Provider;
            if (roles.IsUserInRole(User.Identity.Name, "Admin"))
            {
                return RedirectToAction("Index", "User", new { area = "Admin" });
            }
            if (roles.IsUserInRole(User.Identity.Name, "Teacher"))
            {
                return RedirectToAction("Index", "ThesisisInfo", new { area = "Teachers" });
            }
            ViewBag.Message = "";

            if (roles.IsUserInRole(User.Identity.Name, "Student"))
            {
                return RedirectToAction("HomeStudent", "Home");
            }
            else if (roles.IsUserInRole(User.Identity.Name, "PhdStudent"))
            {
                return RedirectToAction("HomePhdStudent", "Home");
            }

            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your app description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult HomeStudent()
        {
            return View();
        }

        public ActionResult HomePhdStudent()
        {
            return View();
        }
    }
}
