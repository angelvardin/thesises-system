using DissProject.DataLayer;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DissProject.Models;
using System.Data.Entity.Validation;
using System.Diagnostics;

namespace DissProject.Repository
{
    public class UowData : IUowData
    {
        private readonly DbContextImpl context;
        private readonly Dictionary<Type, object> repositories = new Dictionary<Type, object>();

        public UowData()
            : this(new DbContextImpl())
        {
        }

        public UowData(DbContextImpl context)
        {
            this.context = context;
        }

        public IRepository<IndividualPlan> IndividualPlan 
        {
            get { return this.GetRepository<IndividualPlan>(); } 
        }
        public IRepository<UserProfile> UserProfiles
        {
            get { return this.GetRepository<UserProfile>(); }
        }
        public IRepository<Internship> Internship
        {
            get { return this.GetRepository<Internship>(); }
        }

        public IRepository<Person> People
        {
            get { return this.GetRepository<Person>(); }
        }

        public IRepository<PhdStudent> PhdStudents
        {
            get { return this.GetRepository<PhdStudent>(); }
        }

        public IRepository<Teacher> Teachers
        {
            get { return this.GetRepository<Teacher>(); }
        }

        public IRepository<Department> Departments
        {
            get { return this.GetRepository<Department>(); }
        }
        public IRepository<YearWorkPlanApplications> YearWorkPlans
        {
            get { return this.GetRepository<YearWorkPlanApplications>(); }
        }

        public IRepository<Document> Documents
        {
            get { return this.GetRepository<Document>(); }
        }

        private IRepository<T> GetRepository<T>() where T : class
        {
            if (!this.repositories.ContainsKey(typeof(T)))
            {
                var type = typeof(GenericRepository<T>);

                this.repositories.Add( typeof(T), Activator.CreateInstance(type, this.context));
            }

            return (IRepository<T>)this.repositories[typeof(T)];
        }

		public IRepository<Student> Students
		{
			get { return this.GetRepository<Student>(); }
		}


		public IRepository<Thesis> Thesis
		{
			get { return this.GetRepository<Thesis>(); }

		}

		public IRepository<ThesisApplication> ThesisApplications
		{
			get { return this.GetRepository<ThesisApplication>(); }

		}

		public IRepository<ThesisEvaluation> ThesisEvaluations
		{
			get { return this.GetRepository<ThesisEvaluation>(); }

		}

        //public IRepository<Internship> Internships
        //{
        //    get { return this.GetRepository<Internship>(); }

        //}

		public IRepository<InternshipApplication> InternshipApplications
		{
			get { return this.GetRepository<InternshipApplication>(); }

		}


        public int SaveChanges()
        {
            return this.context.SaveChanges();      
        }

        public void Dispose()
        {
            this.context.Dispose();
        }

    }
}
