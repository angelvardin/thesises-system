using System;
using System.Collections.Generic;
using System.Linq;
using System.Transactions;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using DotNetOpenAuth.AspNet;
using Microsoft.Web.WebPages.OAuth;
using WebMatrix.WebData;

using DissProject.Site.Filters;
using DissProject.Site.Models;
using DissProject.Models;
using DissProject.DataLayer;
using DissProject.Repository;

namespace DissProject.Site.Controllers
{
    [Authorize]
    [InitializeSimpleMembership]
    public class AccountController : Controller
    {

        IUowData _db;

        public AccountController()
        {
            this._db = new UowData();
        }


        //[HttpPost]
        //public JsonResult doesUserNameExist(string UserName)
        //{

        //    var user = Membership.GetUser(UserName);

        //    return Json(user == null);
        //}

        //
        // GET: /Account/Login
        [AllowAnonymous]
        public ActionResult Login(string returnUrl)
        {
            ViewBag.ReturnUrl = returnUrl;
            return View();
        }

        //
        // POST: /Account/Login
        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public ActionResult Login(LoginModel model, string returnUrl)
        {
            if (ModelState.IsValid && WebSecurity.Login(model.UserName, model.Password, persistCookie: model.RememberMe))
            {
                var roles = (SimpleRoleProvider)Roles.Provider;
                if (roles.IsUserInRole(model.UserName, "Admin"))
                {
                  return  RedirectToAction("Index", "User", new { area = "Admin" });
                }
                if (roles.IsUserInRole(model.UserName, "Teacher"))
                {
                   return RedirectToAction("Index", "ThesisisInfo", new { area = "Teachers" });
                }
                if (roles.IsUserInRole(model.UserName, "PhdStudent"))
                {
                    return RedirectToAction("HomePhdStudent", "Home");
                }
                if (roles.IsUserInRole(model.UserName, "Student"))
                {
                    return RedirectToAction("HomeStudent", "Home");
                }
                return RedirectToLocal(returnUrl);
            }

            // If we got this far, something failed, redisplay form
            ModelState.AddModelError("", "Потребителското име или парола са некоректни.");
            return View(model);
        }

        //
        // POST: /Account/LogOff
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult LogOff()
        {
            WebSecurity.Logout();

            return RedirectToAction("Index", "Home");
        }

        //
        // GET: /Account/Register
        [AllowAnonymous]
        public ActionResult Register()
        {
            return View();
        }

        //
        // POST: /Account/Register
        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public ActionResult Register(RegisterModel model)
        {
            if (ModelState.IsValid)
            {
                ModelState.AddModelError("", "Възникна грешка. Моля опитайте пак");
                return View(model);
            }
            // If we got this far, something failed, redisplay form
            return View();
        }

        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public ActionResult RegisterStudent(RegisterModel model, StudentInfo student)
        {
            if (ModelState.IsValid)
            {
                try
                {
                    WebSecurity.CreateUserAndAccount(model.UserName, model.Password);

                    var roles = (SimpleRoleProvider)Roles.Provider;
                    roles.AddUsersToRoles(new string[] { model.UserName }, new string[] { model.Role });

                    UserProfile person = _db.UserProfiles.All()
                                .Where(x => x.UserName == model.UserName)
                                .Select(x => x).SingleOrDefault();



                    _db.Students.Add(new Student
                    {
                        PersonId = person.UserId,
                        User = person,
                        Address = model.Address,
                        FirstName = model.FirstName,
                        SecondName = model.SecondName,
                        LastName = model.LastName,
                        PhoneNumber = model.PhoneNumber,
                        FacultyNumber = student.FacultyNumber,
                        FormOfEducation = student.FormOfEducation,
                        GraduationYear = student.GraduationYear,
                        SubjectOfStudies = student.SubjectOfStudies,
                    });
                    person.IsApproved = false;
                    _db.UserProfiles.Update(person);

                    _db.SaveChanges();
                    WebSecurity.Login(model.UserName, model.Password);
                    return RedirectToAction("Index", "Home");
                    

                }
                catch (MembershipCreateUserException e)
                {
                    ModelState.AddModelError("", ErrorCodeToString(e.StatusCode));
                }
            }

            // If we got this far, something failed, redisplay form
            ModelState.AddModelError("", "Възникна грешка. Моля опитайте пак");
            return View("Register", model);
        }


        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public ActionResult RegisterPhdStudent(RegisterModel model, PhdStudentInfo student)
        {
            if (ModelState.IsValid)
            {
                try
                {

                    WebSecurity.CreateUserAndAccount(model.UserName, model.Password);

                    // Roles.AddUserToRole(model.UserName, model.Role);
                    var roles = (SimpleRoleProvider)Roles.Provider;
                    roles.AddUsersToRoles(new string[] { model.UserName }, new string[] { model.Role });

                    UserProfile person = _db.UserProfiles.All()
                                .Where(x => x.UserName == model.UserName)
                                .Select(x => x).SingleOrDefault();
                    Department department = _db.Departments.GetById(student.Department);


                    _db.PhdStudents.Add(new PhdStudent
                    {
                        PersonId = person.UserId,
                        User = person,
                        Address = model.Address,
                        FirstName = model.FirstName,
                        SecondName = model.SecondName,
                        LastName = model.LastName,
                        PhoneNumber = model.PhoneNumber,
                        Code = student.Code,
                        DateOfApproval = student.DateOfApproval,
                        Department = department,
                        Protocol = student.Protocol,
                        FormOfEducation = student.FormOfEducation,
                        SubjectOfStudies = student.SubjectOfStudies,
                    });

                    person.IsApproved = false;
                    _db.UserProfiles.Update(person);

                    _db.SaveChanges();
                    WebSecurity.Login(model.UserName, model.Password);
                    return RedirectToAction("Index", "Home");

                }
                catch (MembershipCreateUserException e)
                {
                    ModelState.AddModelError("", ErrorCodeToString(e.StatusCode));
                }
            }

            // If we got this far, something failed, redisplay form
            ModelState.AddModelError("", "Възникна грешка. Моля опитайте пак");
            return View("Register", model);
        }

        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public ActionResult RegisterTeacher(RegisterModel model, TeacherInfo teacher)
        {
            if (ModelState.IsValid)
            {
                try
                {
                    WebSecurity.CreateUserAndAccount(model.UserName, model.Password);

                    var roles = (SimpleRoleProvider)Roles.Provider;
                    roles.AddUsersToRoles(new string[] { model.UserName }, new string[] { model.Role });

                    UserProfile person = _db.UserProfiles.All()
                                .Where(x => x.UserName == model.UserName)
                                .Select(x => x).SingleOrDefault();
                    Department department = _db.Departments.GetById(teacher.Department);

                    _db.Teachers.Add(new Teacher
                    {
                        PersonId = person.UserId,
                        User = person,
                        Address = model.Address,
                        FirstName = model.FirstName,
                        SecondName = model.SecondName,
                        LastName = model.LastName,
                        PhoneNumber = model.PhoneNumber,
                        DateOfApproval = teacher.DateOfApproval,
                        Department = department,
                        Title = teacher.TeacherTitle,
                    });

                    person.IsApproved = false;
                    _db.UserProfiles.Update(person);

                    _db.SaveChanges();
                    WebSecurity.Login(model.UserName, model.Password);
                    return RedirectToAction("Index", "Home");

                }
                catch (MembershipCreateUserException e)
                {
                    ModelState.AddModelError("", ErrorCodeToString(e.StatusCode));
                }
            }

            // If we got this far, something failed, redisplay form
            ModelState.AddModelError("", "Възникна грешка. Моля опитайте пак");
            return View("Register", model);
        }


        [AllowAnonymous]
        [HttpPost]
        public ActionResult PersonalInfo(string id)
        {
            try
            {
                ViewBag.Departments = new SelectList(_db.Departments.All(), "Id", "Description");
                if (id == "Student")
                {
                    return PartialView("StudentInfo");
                }
                else if (id == "PhdStudent")
                {
                    return PartialView("PhdStudentInfo");
                }
                else if (id == "Teacher")
                {
                    return PartialView("TeacherInfo");
                }
                else
                {
                    ModelState.AddModelError("", "Възникна грешка. Моля опитайте пак");
                    return View("Register");
                }
            }
            catch (Exception)
            {
                ModelState.AddModelError("", "Възникна грешка. Моля опитайте пак");
                return View("Register");
            }

        }



        //
        // GET: /Account/Manage
        public ActionResult Manage(ManageMessageId? message)
        {
            ViewBag.StatusMessage =
                message == ManageMessageId.ChangePasswordSuccess ? "Your password has been changed."
                : message == ManageMessageId.SetPasswordSuccess ? "Your password has been set."
                : message == ManageMessageId.RemoveLoginSuccess ? "The external login was removed."
                : "";
            ViewBag.HasLocalPassword = OAuthWebSecurity.HasLocalAccount(WebSecurity.GetUserId(User.Identity.Name));
            ViewBag.ReturnUrl = Url.Action("Manage");
            return View();
        }

        //
        // POST: /Account/Manage
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Manage(LocalPasswordModel model)
        {
            bool hasLocalAccount = OAuthWebSecurity.HasLocalAccount(WebSecurity.GetUserId(User.Identity.Name));
            ViewBag.HasLocalPassword = hasLocalAccount;
            ViewBag.ReturnUrl = Url.Action("Manage");
            if (hasLocalAccount)
            {
                if (ModelState.IsValid)
                {
                    // ChangePassword will throw an exception rather than return false in certain failure scenarios.
                    bool changePasswordSucceeded;
                    try
                    {
                        changePasswordSucceeded = WebSecurity.ChangePassword(User.Identity.Name, model.OldPassword, model.NewPassword);
                    }
                    catch (Exception)
                    {
                        changePasswordSucceeded = false;
                    }

                    if (changePasswordSucceeded)
                    {
                        return RedirectToAction("Manage", new { Message = ManageMessageId.ChangePasswordSuccess });
                    }
                    else
                    {
                        ModelState.AddModelError("", "The current password is incorrect or the new password is invalid.");
                    }
                }
            }
            else
            {
                // User does not have a local password so remove any validation errors caused by a missing
                // OldPassword field
                ModelState state = ModelState["OldPassword"];
                if (state != null)
                {
                    state.Errors.Clear();
                }

                if (ModelState.IsValid)
                {
                    try
                    {
                        WebSecurity.CreateAccount(User.Identity.Name, model.NewPassword);
                        return RedirectToAction("Manage", new { Message = ManageMessageId.SetPasswordSuccess });
                    }
                    catch (Exception)
                    {
                        ModelState.AddModelError("", String.Format("Unable to create local account. An account with the name \"{0}\" may already exist.", User.Identity.Name));
                    }
                }
            }

            // If we got this far, something failed, redisplay form
            return View(model);
        }










        #region Helpers
        private ActionResult RedirectToLocal(string returnUrl)
        {
            if (Url.IsLocalUrl(returnUrl))
            {
                return Redirect(returnUrl);
            }
            else
            {
                return RedirectToAction("Index", "Home");
            }
        }

        public enum ManageMessageId
        {
            ChangePasswordSuccess,
            SetPasswordSuccess,
            RemoveLoginSuccess,
        }

        internal class ExternalLoginResult : ActionResult
        {
            public ExternalLoginResult(string provider, string returnUrl)
            {
                Provider = provider;
                ReturnUrl = returnUrl;
            }

            public string Provider { get; private set; }
            public string ReturnUrl { get; private set; }

            public override void ExecuteResult(ControllerContext context)
            {
                OAuthWebSecurity.RequestAuthentication(Provider, ReturnUrl);
            }
        }

        private static string ErrorCodeToString(MembershipCreateStatus createStatus)
        {
            // See http://go.microsoft.com/fwlink/?LinkID=177550 for
            // a full list of status codes.
            switch (createStatus)
            {
                case MembershipCreateStatus.DuplicateUserName:
                    return "User name already exists. Please enter a different user name.";

                case MembershipCreateStatus.DuplicateEmail:
                    return "A user name for that e-mail address already exists. Please enter a different e-mail address.";

                case MembershipCreateStatus.InvalidPassword:
                    return "The password provided is invalid. Please enter a valid password value.";

                case MembershipCreateStatus.InvalidEmail:
                    return "The e-mail address provided is invalid. Please check the value and try again.";

                case MembershipCreateStatus.InvalidAnswer:
                    return "The password retrieval answer provided is invalid. Please check the value and try again.";

                case MembershipCreateStatus.InvalidQuestion:
                    return "The password retrieval question provided is invalid. Please check the value and try again.";

                case MembershipCreateStatus.InvalidUserName:
                    return "The user name provided is invalid. Please check the value and try again.";

                case MembershipCreateStatus.ProviderError:
                    return "The authentication provider returned an error. Please verify your entry and try again. If the problem persists, please contact your system administrator.";

                case MembershipCreateStatus.UserRejected:
                    return "The user creation request has been canceled. Please verify your entry and try again. If the problem persists, please contact your system administrator.";

                default:
                    return "An unknown error occurred. Please verify your entry and try again. If the problem persists, please contact your system administrator.";
            }
        }
        #endregion
    }
}
