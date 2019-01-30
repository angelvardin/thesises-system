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
    [DissProject.Site.Filters.InitializeSimpleMembership]
    public class InternshipController : Controller
    {
        IUowData databaseEntities;

        public InternshipController()
        {
            this.databaseEntities = new UowData();

        }

        
        //
        // GET: /Intership/
        [HttpGet]
        public ActionResult Index()
        {
            DISSContext context = DISSContext.Current;
            var entities = context.Entities;
            if (context.CurrentRole == UserRole.Student)
            {
                return View(context.CurrentStudent.CurrentInternship);
            }
            return View();
        }

        // Post Internship index - maybe a bad idea
        // ученик
        [HttpPost]
        public ActionResult EvaluateInternship(Internship internship, FormCollection collection)
        {
            DISSContext context = DISSContext.Current;
            var entities = DISSContext.Current.Entities;
            if( ViewBag.IsInternshipApprovedApplication == false )
            {
                return View("Error");
            }
            try
            {
                Internship datRealInternship = context.CurrentStudent.CurrentInternship;
                datRealInternship.Grade = internship.Grade;
                datRealInternship.InternshipStatus = InternshipStatus.Evaluated;
                entities.Internship.Update(datRealInternship);
                entities.SaveChanges();
                return RedirectToAction("Index");
            }
            catch (Exception e)
            {

                return View("Error");
            }
        }


        // Get: /Internship/AddInternship
        public ActionResult AddInternship()
        {
            ViewBag.People = new SelectList(this.databaseEntities.Teachers.All(), "PersonId", "FirstName").ToList<SelectListItem>();
            ViewBag.PeopleList = this.databaseEntities.People.All();
            //ViewBag.ConsultantId = new List<SelectListItem>();

            return View();

        }

        //
        // POST: /Intership/AddIntership

        [HttpPost]
        public ActionResult AddInternship( InternshipApplication internshipApplication, FormCollection collection )
        {
            ViewBag.People = new SelectList(this.databaseEntities.Teachers.All(), "PersonId", "FirstName").ToList<SelectListItem>();
            try
            {
                Student student = DISSContext.Current.CurrentStudent;
                var entities = DISSContext.Current.Entities;
                if (student == null)
                {
                    return RedirectToAction("Index");
                }

                var consultantId = collection["ConsultantId"];
                string consultantIdWithoutQuote = consultantId.Remove(consultantId.Length - 1);
                int parsed = -1;
                if (Int32.TryParse(consultantIdWithoutQuote, out parsed))
                {
                    Teacher teacher = entities.Teachers.GetById(parsed);
                    if (teacher != null)
                    {
                        internshipApplication.Consultant = teacher;
                    }
                }
                //this.databaseEntities.SaveChanges();

                Internship internship = student.CurrentInternship;
                if (internship == null)
                {
                    internship = new Internship();
                }

                //InternshipApplication.Student = student;
                //InternshipApplication.Internship = internship;

                internship.InternshipApplication = internshipApplication;
                internship.InternshipStatus = InternshipStatus.Applied;
                internship.Id = student.PersonId;
                internship.InternshipManagerId = internshipApplication.Consultant.PersonId;
                internshipApplication.Internship = internship;
                student.CurrentInternship = internship;

                entities.Students.Update(student);
                entities.SaveChanges();

                return RedirectToAction("Index");

            }
            catch (Exception e)
            {
                return View();
            }
        
        }


        // GET: /Internship/AcceptInternship
        //public ActionResult AcceptInternship()
        //{
        //    //ViewBag.Internship = new SelectList(this.databaseEntities.Internship.All(), "Id", "InternshipApplication");
        //    DISSContext context = DISSContext.Current;
        //    var entities = context.Entities;
        //    if (context.CurrentRole == UserRole.Teacher)
        //    {
        //        List<Internship> listOfInternships = new List<Internship>();
        //        foreach (Internship inter in context.CurrentTeacher.ManagerOfInternship)
        //        {
        //            listOfInternships.Add(inter);
        //        }
        //        Internship internshipForAccepting = listOfInternships[0];
        //        return View(internshipForAccepting);
        //    }
        //    return View();
        //}


        // Teacher response to an internship application
        // Post: /Intership/AcceptInternship
        // teqacher
        //[HttpPost]
        //public ActionResult AcceptInternship( Internship internship, FormCollection collection)
        //{
        //    Teacher teacher = DISSContext.Current.CurrentTeacher;
        //    var entities = DISSContext.Current.Entities;
        //    try
        //    {
        //        if (teacher == null)
        //        {
        //            return RedirectToAction("Index");
        //        }
        //        Internship realInternship = entities.Internship.GetById(internship.Id);
        //        if (ViewBag.Evaluated == false)
        //        {
        //            realInternship.InternshipStatus = InternshipStatus.ApprovedApplication;
        //        }
        //        if (ViewBag.Evaluated == true)
        //        {
        //            realInternship.InternshipStatus = InternshipStatus.ApprovedEvaluation;
        //        }
        //        entities.Internship.Update(realInternship);
        //        entities.SaveChanges();

        //        return RedirectToAction("AcceptInternship");
        //    }
        //    catch (Exception e)
        //    {
                
        //        return View();
        //    }

        //}



        //public ActionResult ViewPerson()
        //{
        //    List<Person> persons = databaseEntities.People.All().ToList();
        //    return View(persons);
        //}

      

  

    }
}
