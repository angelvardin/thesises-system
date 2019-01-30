using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DissProject.Models;
using System.Web.Security;
using DissProject.Repository;
//using WebMatrix.WebData.
using System.Web.Security;
using DissProject.Models;
using WebMatrix.WebData;

namespace DissProject
{
    public class DISSContext
    {
        protected const string HTTP_CONTEXT_HCONTEXT_KEY = "DISSContext";

        private IUowData  repository;
        private Person currentPerson;
        private string userName;

        public static object hronoContextLock = new object();

        public static DISSContext hronoContextInstance;

        public Person CurrentPerson
        {
            get
            {
                return currentPerson;
            }
        }

        public Student CurrentStudent
        {
            get
            {
                return this.currentPerson as Student;
            }
        }

        public PhdStudent CurrentPhdStudent
        {
            get
            {
                return this.currentPerson as PhdStudent;
            }
        }

        public Teacher CurrentTeacher
        {
            get
            {
                return this.currentPerson as Teacher;
            }
        }

        public IUowData Entities
        {
            get
            {
                return repository;
            }
        }

        public MembershipUser User
        {
            get
            {
                return Membership.GetUser(userName);
            }
        }

        public UserRole CurrentRole
        {
            get
            {
                var rolesProvider = (SimpleRoleProvider)Roles.Provider;
                string[] roles = rolesProvider.GetRolesForUser(userName);
                string role = roles.SingleOrDefault();
                return UserRoleUtilities.userRoleFromString( role );
            }
        }

        static Person getPersonFromUserName( IUowData entities, string userName )
        {
            UserProfile profile = entities.UserProfiles.All().Where( x => x.UserName == userName ).SingleOrDefault();
            return profile.Person;
        }

        public static DISSContext Current
        {
            get
            {
                //if the context is called from controller
                if (HttpContext.Current != null)
                {
                    HttpContextBase HRContext = new HttpContextWrapper(HttpContext.Current);
                    DISSContext hc = HRContext.Items[HTTP_CONTEXT_HCONTEXT_KEY] as DISSContext;
                    if (hc == null)
                    {
                        lock (DISSContext.hronoContextLock)
                        {
                            hc = HRContext.Items[HTTP_CONTEXT_HCONTEXT_KEY] as DISSContext;
                            if (hc == null)
                            {
                                hc = new DISSContext()
                                {
                                    repository = new UowData(),
                                    //userName = HttpContext.Current.User.Identity.Name,
                                    userName = Membership.GetUser().UserName,
                                };
                                hc.currentPerson = getPersonFromUserName(hc.repository, hc.userName);
                                HRContext.Items[HTTP_CONTEXT_HCONTEXT_KEY] = hc;
                            }
                        }
                    }
                    return hc;
                }
                return hronoContextInstance;
            }
        }

        public static void Initialize(IUowData aRepository, string aUserName)
        {
            hronoContextInstance = new DISSContext()
            {
                repository = new UowData(),
                userName = aUserName
            };
            hronoContextInstance.currentPerson = getPersonFromUserName(hronoContextInstance.repository, hronoContextInstance.userName);
        }
    }
}