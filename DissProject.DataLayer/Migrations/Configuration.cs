using WebMatrix.WebData;

namespace DissProject.DataLayer.Migrations
{
    using DissProject.Models;
    using System;
    using System.Collections.Generic;
    using System.Data.Entity;
    using System.Data.Entity.Migrations;
    using System.Linq;
    using System.Web.Security;

    internal sealed class Configuration : DbMigrationsConfiguration<DissProject.DataLayer.DbContextImpl>
    {
        public Configuration()
        {
            AutomaticMigrationsEnabled = true;
        }

        protected override void Seed(DissProject.DataLayer.DbContextImpl context)
        {
            WebSecurity.InitializeDatabaseConnection("DefaultConnection", "UserProfiles", "UserId", "UserName", autoCreateTables: true);
            
            #region
            var roles = (SimpleRoleProvider)Roles.Provider;
            var membership = (SimpleMembershipProvider)Membership.Provider;

            // initialize accounts, the username is the same as the password
            var initialAccounts = new Dictionary<String, String>()
            {
                // role      username
                { "Admin", "admin" },
                { "Student", "student" },
                { "Teacher", "teacher" },
                { "PhdStudent", "phdStudent" }
            };

            foreach (var role in initialAccounts.Keys)
            {
                String userNameAndPassword = initialAccounts[role];

                if (!roles.RoleExists(role))
                {
                    roles.CreateRole(role);
                }

                if (membership.GetUser(userNameAndPassword, false) == null)
                {
                    membership.CreateUserAndAccount(userNameAndPassword, userNameAndPassword);
                }

                if (!roles.IsUserInRole(userNameAndPassword, role))
                {
                    roles.AddUsersToRoles(new[] { userNameAndPassword }, new[] { role });
                }
            }
            

            // init Departments
            context.Departaments.Add(new Department()
            {
                Description = "ФМИ"
            });
            context.SaveChanges();

            UserProfile student = context.UserProfiles
                                  .Where(x => x.UserName == "student")
                                  .Select(x => x).SingleOrDefault();
            student.IsApproved = true;

            //create Student
            context.Students.AddOrUpdate(new Student()
            {
                PersonId = student.UserId,
                FirstName = "Георги",
                SecondName = "Борисов",
                LastName = "Синеклиев",
                PhoneNumber = "0890539347",
                SubjectOfStudies = "СИ",
                GraduationYear = 4,
                Department = context.Departaments.First(),
                FacultyNumber = 61381,
                Address = "Sofia city",
                User = context.UserProfiles.Find(membership.GetUserId("student")),
                FormOfEducation = FormOfEducation.FullTimeStudy
            });

            context.UserProfiles.AddOrUpdate(student);

            UserProfile teacher = context.UserProfiles
                      .Where(x => x.UserName == "teacher")
                      .Select(x => x).SingleOrDefault();
            teacher.IsApproved = true;

            ////create Teacher
            context.Teachers.Add(new Teacher()
            {
                PersonId = teacher.UserId,
                FirstName = "Учител",
                SecondName = "Учител",
                LastName = "Учител",
                PhoneNumber = "0890539347",
                Department = context.Departaments.First(),
                Address = "Sofia city",
                Title = "Доктор",
                DateOfApproval = DateTime.Today,
                User = context.UserProfiles.Find(membership.GetUserId("teacher"))
            });

            context.UserProfiles.AddOrUpdate(teacher);

            context.SaveChanges();
            #endregion
            
        }
    }
}
