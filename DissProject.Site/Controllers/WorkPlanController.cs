using DissProject.Models;
using DissProject.Repository;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;
using System.Data.Entity;
using System.Web.UI.WebControls;
using WebMatrix.WebData;
using DocumentGeneration;

namespace DissProject.Site.Controllers
{

    [DissProject.Site.Filters.InitializeSimpleMembership]
    public class WorkPlanController : Controller
    {   
        public ActionResult Index()
        {
            int userId = WebSecurity.GetUserId(User.Identity.Name);
            var initial = this.databaseEntities.YearWorkPlans.All()
                .Where(u => u.PhdStudent.PersonId == userId);
            return View(initial);
        }

        public FileResult ExportSingle(int id)
        {
            var s = (from plans in this.databaseEntities.YearWorkPlans.All().Include(x => x.PhdStudent)
                      where plans.Id == id
                     select plans).Single();
            var gen = new YearWorkPlanDocument();
            gen.CreatePackage(Server.MapPath("~/Content/Documents/a.doc"), s);

            return File(Server.MapPath("~/Content/Documents/a.doc"), "application/msword");
        }

        public ActionResult Create()
        {
            return View();
        }

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
                var plan = new YearWorkPlanApplications 
                {PhdStudent = student,
                Manager = student.DirectorOfStudies
                 , Description = collection["Description"]
                 ,
                 DueDate = DateTime.ParseExact(collection["DueDate"], "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture)
                 , FormOfConduct = collection["FormOfConduct"]
                 , FormOfReport = collection["FormOfReport"]
                 , PlanYear = Int32.Parse(collection["PlanYear"])
                 , Title = collection["Title"]
                };
                   entities.YearWorkPlans.Add(plan);

                entities.SaveChanges();
                //return View();

                return RedirectToAction("Index");
            }
            catch(Exception ex)
            {
                return View();
            }
        }

        //
        // GET: /Default1/Edit/5

        public ActionResult Edit(int id)
        {
            int userId = WebSecurity.GetUserId(User.Identity.Name);
            var initial = this.databaseEntities.YearWorkPlans.All()
                .Where(u => u.PhdStudent.PersonId == userId && u.Id==id).Single();
            return View(initial);
        }

        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {

                PhdStudent student = DISSContext.Current.CurrentPhdStudent;
                var entities = DISSContext.Current.Entities;
                entities.YearWorkPlans.Update(
                    new YearWorkPlanApplications()
                    {
                          Id = id
                        , PhdStudent = student
                        , Manager = student.DirectorOfStudies
                        , Description = collection["Description"]
                        , DueDate = DateTime.Parse(collection["DueDate"])
                        , FormOfConduct = collection["FormOfConduct"]
                        , FormOfReport = collection["FormOfReport"]
                        , PlanYear = Int32.Parse(collection["PlanYear"])
                        , Title = collection["Title"]
                    });
                entities.SaveChanges();
                
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        public ActionResult Delete(int id)
        {

            int userId = WebSecurity.GetUserId(User.Identity.Name);
            var initial = this.databaseEntities.YearWorkPlans.All()
                .Where(u => u.PhdStudent.PersonId == userId && u.Id == id).Single();
            return View(initial);
        }

        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                var entities = DISSContext.Current.Entities;
                entities.YearWorkPlans.Delete(id);
                entities.SaveChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
        IUowData databaseEntities;
        //
        // GET: /WorkPlan/

        public WorkPlanController()
        {
            this.databaseEntities = new UowData();
        }
        public ActionResult ExportData()
        {
            GridView gv = new GridView();
            gv.DataSource = this.databaseEntities.YearWorkPlans.All().ToList();
            gv.DataBind();
            Response.ClearContent();
            Response.Buffer = true;
           // Response.AddHeader("content-disposition", "attachment; filename=Marklist.xls");
            Response.ContentType = "application/vnd.ms-excel";
            Response.Charset = "";
            StringWriter sw = new StringWriter();
            HtmlTextWriter htw = new HtmlTextWriter(sw);
            gv.RenderControl(htw);
            Response.Output.Write(sw.ToString());
            Response.Flush();
            Response.End();

            return RedirectToAction("Index");
        }
    }
}
