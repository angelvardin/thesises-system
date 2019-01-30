using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DissProject.Models;

namespace DissProject.Site.Areas.Teachers.Controllers
{
    public class InternshipInfoController : Controller
    {
        //
        // GET: /Teachers/InternshipInfo/

        //public ActionResult Index()
        //{
        //    return View();
        //}

        // Teacher response to an internship application
        // Post: /Intership/AcceptInternship
        // teqacher
        [HttpPost]
        public ActionResult AcceptInternship(Internship internship, FormCollection collection)
        {
            Teacher teacher = DISSContext.Current.CurrentTeacher;
            var entities = DISSContext.Current.Entities;
            try
            {
                if (teacher == null)
                {
                    return RedirectToAction("Error");
                }
                Internship realInternship = entities.Internship.GetById(internship.Id);
                if (realInternship.InternshipStatus == InternshipStatus.Applied)
                {
                    realInternship.InternshipStatus = InternshipStatus.ApprovedApplication;
                }
                else 
                {
                    realInternship.InternshipStatus = InternshipStatus.ApprovedEvaluation;
                }
                entities.Internship.Update(realInternship);
                entities.SaveChanges();

                return RedirectToAction("Index");
            }
            catch (Exception e)
            {

                return View();
            }

        }

        public ActionResult Index()
        {
            //ViewBag.Internship = new SelectList(this.databaseEntities.Internship.All(), "Id", "InternshipApplication");
            DISSContext context = DISSContext.Current;
            var entities = context.Entities;
            Internship internship = entities.Internship.All().Where(i => i.InternshipStatus == InternshipStatus.Applied).SingleOrDefault();

            Internship other = entities.Internship.All().Where(i => i.InternshipStatus == InternshipStatus.Evaluated).SingleOrDefault();
            //if (context.CurrentRole == UserRole.Teacher)
            //{
            //    List<Internship> listOfInternships = new List<Internship>();
            //    foreach (Internship inter in context.CurrentTeacher.ManagerOfInternship)
            //    {
            //        listOfInternships.Add(inter);
            //    }
            //    Internship internshipForAccepting = listOfInternships[0];
            //    return View(internshipForAccepting);
            //}
            Internship result = (internship == null) ?
                            (other == null ? null : other ): internship;
            return View( result );
        }

    }
}
