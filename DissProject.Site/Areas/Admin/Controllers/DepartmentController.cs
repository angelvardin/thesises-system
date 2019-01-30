using DissProject.Models;
using DissProject.Repository;
using DissProject.Site.Areas.Admin.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DissProject.Site.Areas.Admin.Controllers
{
        [Authorize(Roles = "Admin,Teacher")]
    //[DissProject.Site.Filters.IsApproved]
    public class DepartmentController : Controller
    {
        IUowData _db;
        public DepartmentController()
        {
           
            this._db = new UowData();
        }

        //
        // GET: /Admin/Department/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult AllDepartments()
        {
            List<DepartmentViewModel> allDepartments = new List<DepartmentViewModel>();
            allDepartments.AsQueryable<DepartmentViewModel>();
            try
            {
                var departments = _db.Departments.All().ToList();
                foreach (var item in departments)
                {
                    allDepartments.Add(new DepartmentViewModel
                    {
                        DepartmentId = item.Id,
                        Description = item.Description,
                    });
                }
                return View(allDepartments.AsQueryable<DepartmentViewModel>());
            }
            catch (Exception)
            {

                return View(allDepartments.AsQueryable<DepartmentViewModel>());

            }

            
        }
       
 
    }
}
