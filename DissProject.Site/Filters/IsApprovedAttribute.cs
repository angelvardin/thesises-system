using DissProject.Models;
using DissProject.Repository;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;



namespace DissProject.Site.Filters
{


    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Method)]
    public sealed class IsApprovedAttribute : AuthorizeAttribute
    {
        private IUowData _db = new UowData();
        private bool _isAuthorized;
        private bool _isInRole;
        private string RedirectUrl = "~/Home/Index";

        protected override bool AuthorizeCore(System.Web.HttpContextBase httpContext)
        {
            _isInRole = base.AuthorizeCore(httpContext);
            if (_isInRole == false)
            {
                return false;
            }
            string currentUser = HttpContext.Current.User.Identity.Name;
            UserProfile person = _db.UserProfiles.All()
                               .Where(x => x.UserName == currentUser)
                               .Select(x => x).SingleOrDefault();

            if (person.IsApproved == true)
            {
                _isAuthorized = true;
            }
            else
            {
                _isAuthorized = false;
            }

            return _isAuthorized;
        }

        public override void OnAuthorization(AuthorizationContext filterContext)
        {
            base.OnAuthorization(filterContext);

            if (!_isAuthorized)
            {
                filterContext.Controller.TempData.Add("RedirectReason", "Unauthorized");
                //filterContext.RequestContext.HttpContext.Response.Redirect(RedirectUrl);
            }
        }
    }
}