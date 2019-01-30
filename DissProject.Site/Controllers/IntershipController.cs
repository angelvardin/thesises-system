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
    public class IntershipController : Controller
    {
        IUowData databaseEntities;

        public IntershipController()
        {
            this.databaseEntities = new UowData();

        }

        
        //
        // GET: /Intership/

        public ActionResult Index(){

            return View();
        }


        //
        // POST: /Intership/AddIntership
        
        public ActionResult AddIntership()
        {
           // ViewBag.People = new SelectList(this.databaseEntities.People.All(), "PersonId", "FirstName");
            return View();

        
        }



        //public ActionResult ViewPerson()
        //{
        //    List<Person> persons = databaseEntities.People.All().ToList();
        //    return View(persons);
        //}

      

  

    }
}
