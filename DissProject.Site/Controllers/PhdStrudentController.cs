using DissProject.Models;
using DissProject.Repository;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DissProject.Site.Controllers
{
    [DissProject.Site.Filters.InitializeSimpleMembership]
    public class PhdStudentController : Controller
    {
        IUowData databaseEntities;

        public PhdStudentController()
        {
            this.databaseEntities = new UowData();
        }
        //
        // GET: /PhdStrudent/

        public ActionResult Index()
        {
            if (User.IsInRole(UserRoleUtilities.userRoleToString(UserRole.PhdStudent)))
            {
                return View();
            }
            return RedirectToAction("Index", "Home");
        }

        public ActionResult Inquiries()
        {
            return View();
        }

        [HttpPost]
        public JsonResult GetAllGraduatePhdsInRange(string fromDate, string toDate)
        {
            var startDate = DateTime.ParseExact(fromDate, "dd.MM.yyyy", CultureInfo.InvariantCulture);
            var endDate = DateTime.ParseExact(toDate, "dd.MM.yyyy", CultureInfo.InvariantCulture);

            if (startDate == endDate)
                endDate = endDate.AddHours(23).AddMinutes(59).AddSeconds(59);

            var result = (from phdStudents in this.databaseEntities.PhdStudents.All()
                          where phdStudents.DateOfEarnedTitle != null
                            && startDate.CompareTo(phdStudents.DateOfEarnedTitle.Value) < 0
                            && endDate.CompareTo(phdStudents.DateOfEarnedTitle.Value) > 0
                        select phdStudents).ToArray();

            return Json(
                result.Select(r => new
                {
                    Name = String.Format("{0} {1} {2}", r.FirstName, r.SecondName, r.LastName),
                    Date = r.DateOfEarnedTitle.Value.ToString("dd.MM.yyyy")
                }));
        }

        public JsonResult GetAllNotGraduatePhdsInRange(string fromDate, string toDate)
        {
            var startDate = DateTime.ParseExact(fromDate, "dd.MM.yyyy", CultureInfo.InvariantCulture);
            var endDate = DateTime.ParseExact(toDate, "dd.MM.yyyy", CultureInfo.InvariantCulture);

            if (startDate == endDate)
                endDate = endDate.AddHours(23).AddMinutes(59).AddSeconds(59);

            var result = (from phdStudents in this.databaseEntities.PhdStudents.All()
                          where phdStudents.DateOfEarnedTitle == null
                            && phdStudents.DateOfApproval.CompareTo(startDate) > 0
                            && phdStudents.DateOfApproval.CompareTo(endDate) < 0
                          select phdStudents).ToArray();

            return Json(
                result.Select(r => new
                {
                    Name = String.Format("{0} {1} {2}", r.FirstName, r.SecondName, r.LastName),
                    Date = r.DateOfApproval.ToString("dd.MM.yyyy")
                }));
        }

        public JsonResult GetAllGraduatePhdsInRangeWithManager(string fromDate, string toDate, int teacherId)
        {
            var startDate = DateTime.ParseExact(fromDate, "dd.MM.yyyy", CultureInfo.InvariantCulture);
            var endDate = DateTime.ParseExact(toDate, "dd.MM.yyyy", CultureInfo.InvariantCulture);

            if (startDate == endDate)
                endDate = endDate.AddHours(23).AddMinutes(59).AddSeconds(59);

            var result = (from phdStudents in this.databaseEntities.PhdStudents.All()
                          where phdStudents.DateOfEarnedTitle != null
                            && startDate.CompareTo(phdStudents.DateOfEarnedTitle.Value) < 0
                            && endDate.CompareTo(phdStudents.DateOfEarnedTitle.Value) > 0
                            && phdStudents.IndividualPlan.Manager.PersonId == teacherId
                          select phdStudents).ToArray();

            return Json(result.Select(r=> new { 
                Name = r.AllNames,
                Date = r.DateOfEarnedTitle.Value.ToString("dd.MM.yyyy"),
                TeacherName = r.IndividualPlan.Manager.AllNames,
            }));
        }

        public JsonResult GetAllNotGraduatePhdsInRangeWithManager(string fromDate, string toDate, int teacherId)
        {
            var startDate = DateTime.ParseExact(fromDate, "dd.MM.yyyy", CultureInfo.InvariantCulture);
            var endDate = DateTime.ParseExact(toDate, "dd.MM.yyyy", CultureInfo.InvariantCulture);

            if (startDate == endDate)
                endDate = endDate.AddHours(23).AddMinutes(59).AddSeconds(59);

            var result = (from phdStudents in this.databaseEntities.PhdStudents.All()
                          where phdStudents.DateOfEarnedTitle != null
                            && startDate.CompareTo(phdStudents.DateOfApproval) < 0
                            && endDate.CompareTo(phdStudents.DateOfApproval) > 0
                          select phdStudents).ToArray();

            return Json(result.Select(r => new
            {
                Name = r.AllNames,
                Date = r.DateOfApproval.ToString("dd.MM.yyyy"),
                TeacherName = r.IndividualPlan.Manager.AllNames,
            }));
        }

        [HttpPost]
        public JsonResult GetAllTeachers()
        {
            var result = this.databaseEntities.Teachers.All().ToList();

            return Json(result.Select(r => new { 
                Name = r.AllNames,
                Id = r.PersonId }
            ));
        }
    }
}
