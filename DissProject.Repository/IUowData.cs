using DissProject.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DissProject.Repository
{
    public interface IUowData : IDisposable
    {
        IRepository<UserProfile> UserProfiles { get; }

        IRepository<Person> People { get; }

        IRepository<PhdStudent> PhdStudents { get; }

        IRepository<Teacher> Teachers { get; }

        IRepository<Department> Departments { get; }

        IRepository<Student> Students { get; }

        IRepository<YearWorkPlanApplications> YearWorkPlans  { get; }

        IRepository<IndividualPlan> IndividualPlan { get; }

        IRepository<Thesis> Thesis { get; }

        IRepository<ThesisApplication> ThesisApplications { get; }

        IRepository<ThesisEvaluation> ThesisEvaluations { get; }

        IRepository<Document> Documents { get; }

       // IRepository<Internship> Internships { get; }

        IRepository<InternshipApplication> InternshipApplications { get; }

        IRepository<Internship> Internship { get; }

        int SaveChanges();
    }
}
