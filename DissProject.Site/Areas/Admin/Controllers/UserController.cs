using DissProject.Models;
using DissProject.Repository;
using DissProject.Site.Areas.Admin.Models;
using DissProject.Site.Filters;
using DissProject.Site.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using WebMatrix.WebData;

namespace DissProject.Site.Areas.Admin.Controllers
{

    [Authorize(Roles = "Admin,Teacher")]
    //[DissProject.Site.Filters.IsApproved]
    public class UserController : Controller
    {
        IUowData _db;

        public UserController()
        {
            if (!WebSecurity.Initialized)
            {
                WebSecurity.InitializeDatabaseConnection("DefaultConnection", "UserProfiles", "UserId", "UserName", autoCreateTables: true);
            }
            this._db = new UowData();
        }
        //
        // GET: /Admin/User/


        public ActionResult Index()
        {
            if (!WebSecurity.Initialized)
            {
                WebSecurity.InitializeDatabaseConnection("DefaultConnection", "UserProfiles", "UserId", "UserName", autoCreateTables: true);
            }
            return View();
        }

        public ActionResult AllUsers()
        {
            if (!WebSecurity.Initialized)
            {
                WebSecurity.InitializeDatabaseConnection("DefaultConnection", "UserProfiles", "UserId", "UserName", autoCreateTables: true);
            }
            return View();
        }

        public ActionResult UserList()
        {
            List<UnapproveUser> unapprovedUsers = new List<UnapproveUser>();

            var roles = (SimpleRoleProvider)Roles.Provider;

            var allusers = _db.UserProfiles.All().ToList();
            try
            {
                foreach (var item in allusers)
                {

                    if (item.IsApproved == false)
                    {
                        string role = "";
                        if (roles.IsUserInRole(item.UserName, "Administrator"))
                        {
                            role = "Administrator";
                        }
                        else if (roles.IsUserInRole(item.UserName, "Teacher"))
                        {
                            role = "Преподавател";
                        }
                        else if (roles.IsUserInRole(item.UserName, "Student"))
                        {
                            role = "Студент";
                        }
                        else if (roles.IsUserInRole(item.UserName, "PhdStudent"))
                        {
                            role = "Докторант";
                        }
                        else
                        {
                            role = "Грешка";
                        }
                        Person person = item.Person;
                        unapprovedUsers.Add(new UnapproveUser
                        {
                            UserId = item.UserId,
                            UserName = item.UserName,
                            FirstName = person.FirstName,
                            LastName = person.LastName,
                            Role = role,
                        });

                    }
                }
            }
            catch (Exception)
            {
                
                return View(unapprovedUsers.AsQueryable<UnapproveUser>());
            }


            return View(unapprovedUsers.AsQueryable<UnapproveUser>());
        }

        public ActionResult AllUserList()
        {
            List<UnapproveUser> unapprovedUsers = new List<UnapproveUser>();

            var roles = (SimpleRoleProvider)Roles.Provider;

            var allusers = _db.UserProfiles.All().ToList();
            try
            {
                foreach (var item in allusers)
                {


                    string role = "";
                    if (roles.IsUserInRole(item.UserName, "Administrator"))
                    {

                        role = "Administrator";
                        continue;
                    }
                    else if (roles.IsUserInRole(item.UserName, "Teacher"))
                    {
                        role = "Преподавател";
                    }
                    else if (roles.IsUserInRole(item.UserName, "Student"))
                    {
                        role = "Студент";
                    }
                    else if (roles.IsUserInRole(item.UserName, "PhdStudent"))
                    {
                        role = "Докторант";
                    }
                    else
                    {
                        role = "error";
                        continue;
                    }

                    Person person = item.Person;
                    if (person == null)
                    {
                        continue;
                    }

                    unapprovedUsers.Add(new UnapproveUser
                    {
                        UserId = item.UserId,
                        UserName = item.UserName,
                        FirstName = person.FirstName,
                        LastName = person.LastName,
                        Role = role,
                    });


                }
            }
            catch (Exception)
            {

                ModelState.AddModelError("", "Възникна грешка. Моля опитайте пак");
                return View(unapprovedUsers.AsQueryable<UnapproveUser>());
            }


            return View(unapprovedUsers.AsQueryable<UnapproveUser>());
        }

        [HttpPost]
        public string SendEmail(int user)
        {
            try
            {
                UserProfile person = _db.UserProfiles.All()
                    .Where(x => x.UserId == user)
                    .Select(x => x).SingleOrDefault();
                if (person == null)
                {
                    return "fail";
                }
                //Накрая ще активираме пращането на мейл. Предложете имена за обща поща! :)
                //var client = new SmtpClient("smtp.gmail.com", 587)
                //{
                //    Credentials = new NetworkCredential("dissproject.fmi@gmail.com", "Pi$$ProjecT.Test"),
                //    EnableSsl = true
                //};
                //client.Send("dissproject.fmi@gmail.com", person.Person.Address , "Одобрение", "Вие бяхте одобрени");

            }
            catch (Exception)
            {
                return "fail";
            }
            return "success";
        }
  
        [HttpPost]
        public string ApproveUser(int user)
        {
            try
            {
                UserProfile person = _db.UserProfiles.All()
                    .Where(x => x.UserId == user)
                    .Select(x => x).SingleOrDefault();
                if (person == null)
                {
                    return "fail";
                }

                person.IsApproved = true;
                _db.UserProfiles.Update(person);
                _db.SaveChanges();
            }
            catch (Exception)
            {
                return "fail";
            }
            return "success";
        }

        [HttpPost]
        public string DeleteUser(int user)
        {
            try
            {
                var roles = (SimpleRoleProvider)Roles.Provider;
                UserProfile person = _db.UserProfiles.All()
                    .Where(x => x.UserId == user)
                    .Select(x => x).SingleOrDefault();
                if (person == null)
                {
                    return "fail";
                }

                if (person.UserName == User.Identity.Name)
                {
                    return "deleteyourself";
                }

                bool isAdmin = roles.IsUserInRole(person.UserName, "Administrator");
                if (isAdmin)
                {
                    return "noadmin";
                }

                if (roles.IsUserInRole(person.UserName, "Administrator"))
                {
                    string name = person.UserName;
                    Teacher teacher = _db.Teachers.GetById(person.UserId);
                    _db.Teachers.Delete(teacher);
                    roles.RemoveUsersFromRoles(new string[] { name }, new string[] { "Teacher" });
                    _db.SaveChanges();
                    ((SimpleMembershipProvider)Membership.Provider).DeleteAccount(name);
                    ((SimpleMembershipProvider)Membership.Provider).DeleteUser(name, true);

                }
                else if (roles.IsUserInRole(person.UserName, "Teacher"))
                {
                    string name = person.UserName;
                    Teacher teacher = _db.Teachers.GetById(person.UserId);
                    _db.Teachers.Delete(teacher);
                    roles.RemoveUsersFromRoles(new string[] { name }, new string[] { "Teacher" });
                    _db.SaveChanges();
                    ((SimpleMembershipProvider)Membership.Provider).DeleteAccount(name);
                    ((SimpleMembershipProvider)Membership.Provider).DeleteUser(name, true);
                }
                else if (roles.IsUserInRole(person.UserName, "PhdStudent"))
                {
                    string name = person.UserName;
                    PhdStudent teacher = _db.PhdStudents.GetById(person.UserId);
                    _db.PhdStudents.Delete(teacher);
                    roles.RemoveUsersFromRoles(new string[] { name }, new string[] { "PhdStudent" });
                    _db.SaveChanges();
                    ((SimpleMembershipProvider)Membership.Provider).DeleteAccount(name);
                    ((SimpleMembershipProvider)Membership.Provider).DeleteUser(name, true);
                }
                else if (roles.IsUserInRole(person.UserName, "Student"))
                {
                    string name = person.UserName;
                    Student teacher = _db.Students.GetById(person.UserId);
                    _db.Students.Delete(teacher);
                    roles.RemoveUsersFromRoles(new string[] { name }, new string[] { "Student" });
                    _db.SaveChanges();
                    ((SimpleMembershipProvider)Membership.Provider).DeleteAccount(name);
                    ((SimpleMembershipProvider)Membership.Provider).DeleteUser(name, true);
                }
                else
                {
                    return "fail";
                }


            }
            catch (Exception)
            {
                return "fail";
            }
            return "success";
        }

        public ActionResult Details(int userId)
        {
            if (!WebSecurity.Initialized)
            {
                WebSecurity.InitializeDatabaseConnection("DefaultConnection", "UserProfiles", "UserId", "UserName", autoCreateTables: true);
            }
            var roles = (SimpleRoleProvider)Roles.Provider;

            UserProfile person = _db.UserProfiles.All()
                    .Where(x => x.UserId == userId)
                    .Select(x => x).SingleOrDefault();

            if (roles.IsUserInRole(person.UserName, "Teacher"))
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
                return View("TeacherDetails", model);
            }
            else if (roles.IsUserInRole(person.UserName, "Student"))
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
                return View("StudentDetails", model);
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
                    Protocol = phdStudent.Protocol,
                };
                return View("PhdStudentDetails", model);
            }
            else
            {
                ModelState.AddModelError("", "Възникна грешка. Моля опитайте пак");
            }


            return View("Index");
        }

        //
        // GET: /Admin/User/Create

        public ActionResult Create()
        {
            return View();
        }

        //
        // POST: /Admin/User/Create

        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
                // TODO: Add insert logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /Admin/User/Edit/5

        public ActionResult Edit(int id)
        {
            return View();
        }

        //
        // POST: /Admin/User/Edit/5

        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /Admin/User/Delete/5

        public ActionResult Delete(int id)
        {
            return View();
        }

        //
        // POST: /Admin/User/Delete/5

        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
    }
}
