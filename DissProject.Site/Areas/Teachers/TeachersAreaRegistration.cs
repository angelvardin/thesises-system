using System.Web.Mvc;

namespace DissProject.Site.Areas.Teachers
{
    public class TeachersAreaRegistration : AreaRegistration
    {
        public override string AreaName
        {
            get
            {
                return "Teachers";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context)
        {
            context.MapRoute(
                "Teachers_default",
                "Teachers/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}
