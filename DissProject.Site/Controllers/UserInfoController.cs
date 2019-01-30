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

namespace DissProject.Site.Controllers
{
    [InitializeSimpleMembership]
    public class UserInfoController : Controller
    {
        IUowData _db;

        public UserInfoController()
        {
            this._db = new UowData();
        }
        //
        // GET: /User/Edit/5

        public ActionResult Edit()
        {
            UserProfile person = _db.UserProfiles.All()
                                .Where(x => x.UserName == User.Identity.Name)
                                .Select(x => x).SingleOrDefault();

    


            var roles = (SimpleRoleProvider)Roles.Provider;

            ViewBag.Departments = new SelectList(_db.Departments.All(), "Id", "Description");
            if (roles.IsUserInRole(person.UserName, "Student"))
            {
                Student student = _db.Students.GetById(person.UserId);
                EditStudentInfo model = new EditStudentInfo() 
                {
                    FirstName = student.FirstName,
                    LastName = student.LastName,
                    SecondName = student.SecondName,
                    Address = student.Address,
                    PhoneNumber = student.PhoneNumber,
                    FacultyNumber = student.FacultyNumber,
                    SubjectOfStudies = student.SubjectOfStudies,
                    FormOfEducation = student.FormOfEducation,
                    GraduationYear = student.GraduationYear,  
                };
                return View("EditStudent", model);
            }
            else if (roles.IsUserInRole(person.UserName, "Teacher"))
            {
                Teacher teacher = _db.Teachers.GetById(person.UserId);

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
            else if (roles.IsUserInRole(person.UserName, "PhdStudent"))
            {
                PhdStudent phdStudent = _db.PhdStudents.GetById(person.UserId);
                EditPhdStudentInfo model = new EditPhdStudentInfo()
                {
                    FirstName = phdStudent.FirstName,
                    LastName = phdStudent.LastName,
                    SecondName = phdStudent.SecondName,
                    Address = phdStudent.Address,
                    PhoneNumber = phdStudent.PhoneNumber,
                    SubjectOfStudies = phdStudent.SubjectOfStudies,
                    FormOfEducation = phdStudent.FormOfEducation,
                    Code = phdStudent.Code,
                    DateOfApproval = phdStudent.DateOfApproval,
                    Department = phdStudent.Department.Id,
                    Protocol  = phdStudent.Protocol,
                };
                return View("EditPhdStudent", model);
            }




            return RedirectToAction("Index", "Home");
        }

        //
        // POST: /User/Edit/5

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

                    return RedirectToAction("Index", "Home");
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

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult EditStudentInfo(EditStudentInfo model)
        {
            if (ModelState.IsValid)
            {
                try
                {
                    // TODO: Add update logic here
                    UserProfile person = _db.UserProfiles.All()
                                    .Where(x => x.UserName == User.Identity.Name)
                                    .Select(x => x).SingleOrDefault();
                    Student student = _db.Students.GetById(person.UserId);
                    student.Address = model.Address;
                    student.FirstName = model.FirstName;
                    student.LastName = model.LastName;
                    student.PhoneNumber = model.PhoneNumber;
                    student.SecondName = model.SecondName;
                    student.FacultyNumber = model.FacultyNumber;
                    student.FormOfEducation = model.FormOfEducation;
                    student.GraduationYear = model.GraduationYear;
                    student.SubjectOfStudies = model.SubjectOfStudies;
                    _db.Students.Update(student);
                    _db.SaveChanges();

                    return RedirectToAction("Index", "Home");
                }
                catch
                {
                    ModelState.AddModelError("", "Възникна грешка. Моля опитайте пак");
                    return View("EditStudent", model);
                }
            }

            ModelState.AddModelError("", "Възникна грешка. Моля опитайте пак");
            return View("EditStudent", model);
           
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult EditPhdStudentInfo(EditPhdStudentInfo model)
        {
            if (ModelState.IsValid)
            {
                try
                {
                    // TODO: Add update logic here
                    UserProfile person = _db.UserProfiles.All()
                                    .Where(x => x.UserName == User.Identity.Name)
                                    .Select(x => x).SingleOrDefault();
                    PhdStudent phdStudent = _db.PhdStudents.GetById(person.UserId);
                    phdStudent.Address = model.Address;
                    phdStudent.FirstName = model.FirstName;
                    phdStudent.LastName = model.LastName;
                    phdStudent.PhoneNumber = model.PhoneNumber;
                    phdStudent.SecondName = model.SecondName;
                    phdStudent.FormOfEducation = model.FormOfEducation;
                    phdStudent.Protocol = model.Protocol;
                    phdStudent.Code = model.Code;
                    phdStudent.Department = _db.Departments.GetById(model.Department);
                    phdStudent.DateOfApproval = model.DateOfApproval;
                    phdStudent.SubjectOfStudies = model.SubjectOfStudies;

                    _db.PhdStudents.Update(phdStudent);
                    _db.SaveChanges();

                    return RedirectToAction("Index", "Home");
                }
                catch
                {
                    ModelState.AddModelError("", "Възникна грешка. Моля опитайте пак");
                    return View("EditPhdStudent", model);
                }
            }
            ViewBag.Departments = new SelectList(_db.Departments.All(), "Id", "Description");
            ModelState.AddModelError("", "Възникна грешка. Моля опитайте пак");
            return View("EditPhdStudent", model);

        }

        ////
        //// GET: /User/Delete/5

        //public ActionResult Delete(int id)
        //{
        //    return View();
        //}

        ////
        //// POST: /User/Delete/5

        //[HttpPost]
        //public ActionResult Delete(int id, FormCollection collection)
        //{
        //    try
        //    {
        //        // TODO: Add delete logic here

        //        return RedirectToAction("Index");
        //    }
        //    catch
        //    {
        //        return View();
        //    }
        //}
    }
}
