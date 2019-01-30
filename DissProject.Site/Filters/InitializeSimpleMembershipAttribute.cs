using System;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Threading;
using System.Web.Mvc;
using WebMatrix.WebData;
using DissProject.Site.Models;
using DissProject.DataLayer;
using System.Collections.Generic;
using System.Web.Security;

namespace DissProject.Site.Filters
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Method, AllowMultiple = false, Inherited = true)]
    public sealed class InitializeSimpleMembershipAttribute : ActionFilterAttribute
    {
        private static SimpleMembershipInitializer _initializer;
        private static object _initializerLock = new object();
        private static bool _isInitialized;

        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            // Ensure ASP.NET Simple Membership is initialized only once per app start
            LazyInitializer.EnsureInitialized(ref _initializer, ref _isInitialized, ref _initializerLock);
        }

        private class SimpleMembershipInitializer
        {
            public SimpleMembershipInitializer()
            {
                Database.SetInitializer<DbContextImpl>(null);
                if ( !WebSecurity.Initialized )
                {
                    WebSecurity.InitializeDatabaseConnection("DefaultConnection", "UserProfiles", "UserId", "UserName", autoCreateTables: true);
                }
            // moved initialization to data layer:)
            //    try
            //    {
            //        using (var context = new DbContextImpl())
            //        {
            //            if (!context.Database.Exists())
            //            {
            //                // Create the SimpleMembership database without Entity Framework migration schema
            //                ((IObjectContextAdapter)context).ObjectContext.CreateDatabase();
            //            }
            //        }
            //        if (!WebSecurity.Initialized)
            //        {
            //            WebSecurity.InitializeDatabaseConnection("DefaultConnection", "UserProfiles", "UserId", "UserName", autoCreateTables: true);

            //            // custom code for creating users here.

            //            var roles = (SimpleRoleProvider)Roles.Provider;
            //            var membership = (SimpleMembershipProvider)Membership.Provider;

            //            // initialize accounts, the username is the same as the password
            //            var initialAccounts = new Dictionary<String, String>()
            //            {
            //                // role      username
            //                { "Admin", "admin" },
            //                { "Student", "student" },
            //                { "Teacher", "teacher" },
            //                { "PhdStudent", "phdStudent" }
            //            };

            //            foreach (var role in initialAccounts.Keys)
            //            {
            //                String userNameAndPassword = initialAccounts[role];

            //                if (!roles.RoleExists(role))
            //                {
            //                    roles.CreateRole(role);
            //                }

            //                if (membership.GetUser(userNameAndPassword, false) == null)
            //                {
            //                    membership.CreateUserAndAccount(userNameAndPassword, userNameAndPassword);
            //                }

            //                if (!roles.IsUserInRole(userNameAndPassword, role))
            //                {
            //                    roles.AddUsersToRoles(new[] { userNameAndPassword }, new[] { role });
            //                }
            //            }
            //        }

            //    }
            //    catch (Exception ex)
            //    {
            //        throw new InvalidOperationException("The ASP.NET Simple Membership database could not be initialized. For more information, please see http://go.microsoft.com/fwlink/?LinkId=256588", ex);
            //    }
            }
        }
    }
}
