using DissProject.Models;
using DissProject.Repository;
using DocumentGeneration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebMatrix.WebData;
using System.Data.Entity;


namespace DissProject.Site.Controllers
{
    [DissProject.Site.Filters.InitializeSimpleMembership]
    public class IndividualPlanController : Controller
    {
        IUowData databaseEntities;
        //
        // GET: /WorkPlan/

        public IndividualPlanController()
        {
            this.databaseEntities = new UowData();
        }
        //
        // GET: /IndividualPlan/

        public ActionResult Index()
        {
            int userId = WebSecurity.GetUserId(User.Identity.Name);
            var initial = this.databaseEntities.IndividualPlan.All()
                .Where(u => u.PhdStudent.PersonId == userId);
            return View(initial);
        }

        //
        // GET: /IndividualPlan/Details/5

        public ActionResult Details(int id)
        {
            return View();
        }

        //
        // GET: /IndividualPlan/Create

        public ActionResult Create()
        {
            return View();
        }

        //
        // POST: /IndividualPlan/Create

        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {

                PhdStudent student = DISSContext.Current.CurrentPhdStudent;
                var entities = DISSContext.Current.Entities;
                if (student == null)
                {
                    return RedirectToAction("Index");
                }
                var plan = new IndividualPlan
                {
                    PhdStudent = student,
                    Manager = student.DirectorOfStudies
                 ,
                    GratuationDate = DateTime.ParseExact(collection["GratuationDate"], "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture)
                 ,
                    PhdThesisTitle = collection["PhdThesisTitle"]
                 ,
                    FacultyProtocol = collection["FacultyProtocol"]
                 ,
                    Specialty = collection["Specialty"]
                 
                };
                student.IndividualPlan = plan;
                entities.PhdStudents.Update(student);

                entities.SaveChanges();
                //return View();

                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                return View();
            }
        }

        //
        // GET: /IndividualPlan/Edit/5

        public ActionResult Edit(int id)
        {
            int userId = WebSecurity.GetUserId(User.Identity.Name);
            var initial = this.databaseEntities.IndividualPlan.All()
                .Where(u => u.PhdStudent.PersonId == userId && u.Id == id).Single();
            return View(initial);
        }

        //
        // POST: /IndividualPlan/Edit/5

        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                PhdStudent student  = DISSContext.Current.CurrentPhdStudent;
                var entities        = DISSContext.Current.Entities;
                
                  var plan =  new IndividualPlan
                    {
                          Id = id
                        , PhdStudent = student
                        , Manager = student.DirectorOfStudies
                        , GratuationDate =DateTime.Parse( collection["GratuationDate"])
                        , PhdThesisTitle = collection["PhdThesisTitle"]
                        , FacultyProtocol = collection["FacultyProtocol"]
                        , Specialty = collection["Specialty"]
                               
                    };

                student.IndividualPlan = plan;
                entities.SaveChanges();
                
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        //
        // GET: /IndividualPlan/Delete/5

        public ActionResult Delete(int id)
        {

            int userId = WebSecurity.GetUserId(User.Identity.Name);
            var initial = this.databaseEntities.IndividualPlan.All()
                .Where(u => u.PhdStudent.PersonId == userId && u.Id == id).Single();
            return View(initial);
        }

        //
        // POST: /IndividualPlan/Delete/5

        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                var entities = DISSContext.Current.Entities;
                entities.IndividualPlan.Delete(id);
                entities.SaveChanges();

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
        public FileResult ExportData()
        {
            var student = DISSContext.Current.Entities;
            var sa = (from plan in student.IndividualPlan.All().Include(x => x.PhdStudent)
                                          select plan).Single();
            var s = new GeneratedClass();
            s.CreatePackage(Server.MapPath("~/Content/Documents/a.doc"), sa);

            return File(Server.MapPath("~/Content/Documents/a.doc"), "application/msword");
        }
    }
}
