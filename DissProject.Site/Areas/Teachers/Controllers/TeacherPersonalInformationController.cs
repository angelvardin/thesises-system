using DissProject.Models;
using DissProject.Repository;
using DissProject.Site.Filters;
using DissProject.Site.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using WebMatrix.WebData;

namespace DissProject.Site.Areas.Teachers.Controllers
{
    [InitializeSimpleMembership]
    [IsApproved]
    public class TeacherPersonalInformationController : Controller
    {
        IUowData _db;

        public TeacherPersonalInformationController()
        {
            this._db = new UowData();
        }

        //
        // GET: /Teachers/TeacherPersonalInformation/Edit/
        public ActionResult Edit()
        {
            UserProfile person = _db.UserProfiles.All()
                                .Where(x => x.UserName == User.Identity.Name)
                                .Select(x => x).SingleOrDefault();

            var roles = (SimpleRoleProvider)Roles.Provider;

            ViewBag.Departments = new SelectList(_db.Departments.All(), "Id", "Description");
            Teacher teacher = _db.Teachers.GetById(person.UserId);
            if (teacher == null)
            {
                return RedirectToAction("Index", "ThesisisInfo");
            }

            EditTeacherInfo model = new EditTeacherInfo()
            {
                FirstName = teacher.FirstName,
                LastName = teacher.LastName,
                SecondName = teacher.SecondName,
                Address = teacher.Address,
                PhoneNumber = teacher.PhoneNumber,
                TeacherTitle = teacher.Title,
                Department = teacher.Department.Id,
                DateOfApproval = teacher.DateOfApproval,
            };
            return View("EditTeacher", model);
        }

        //
        // POST: /Teachers/TeacherPersonalInformation/Edit/5

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult EditTeacherInfo(EditTeacherInfo model)
        {
            if (ModelState.IsValid)
            {
                try
                {
                    // TODO: Add update logic here
                    UserProfile person = _db.UserProfiles.All()
                                    .Where(x => x.UserName == User.Identity.Name)
                                    .Select(x => x).SingleOrDefault();
                    Teacher teacher = _db.Teachers.GetById(person.UserId);
                    teacher.Address = model.Address;
                    teacher.DateOfApproval = model.DateOfApproval;
                    teacher.Department = _db.Departments.GetById(model.Department);
                    teacher.FirstName = model.FirstName;
                    teacher.LastName = model.LastName;
                    teacher.PhoneNumber = model.PhoneNumber;
                    teacher.SecondName = model.SecondName;
                    teacher.Title = model.TeacherTitle;
                    _db.Teachers.Update(teacher);
                    _db.SaveChanges();

                    return RedirectToAction("Index", "ThesisisInfo");
                }
                catch
                {

                    ViewBag.Departments = new SelectList(_db.Departments.All(), "Id", "Description");
                    ModelState.AddModelError("", "Възникна грешка. Моля опитайте пак");
                    return View("EditTeacher", model);
                }
            }
            ViewBag.Departments = new SelectList(_db.Departments.All(), "Id", "Description");
            ModelState.AddModelError("", "Възникна грешка. Моля опитайте пак");
            return View("EditTeacher", model);
        }
    }
}
