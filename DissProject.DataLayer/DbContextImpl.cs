using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Entity;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using DissProject.Models;
using System.Data.Entity.ModelConfiguration.Conventions;

namespace DissProject.DataLayer
{
    public class DbContextImpl : DbContext
    {
        public DbContextImpl()
            : base("DefaultConnection")
        {

        }

        public DbSet<UserProfile> UserProfiles { get; set; }
        public DbSet<Person> People { get; set; }
        public DbSet<Student> Students { get; set; }
        public DbSet<PhdStudent> PhdStudents { get; set; }
        public DbSet<Teacher> Teachers { get; set; }
        public DbSet<Department> Departaments { get; set; }
        public DbSet<YearWorkPlanApplications> YearPlans { get; set; }
        public DbSet<IndividualPlan> IndividualPlan { get; set; }
        public DbSet<Document> Documents { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            
            modelBuilder.Conventions.Remove<ManyToManyCascadeDeleteConvention>();

            modelBuilder.Entity<Person>()
                .HasRequired(pi => pi.User).WithRequiredDependent(x => x.Person);

            modelBuilder.Entity<Student>()
                .HasOptional(s => s.CurrentThesis).WithRequired(s => s.Student);

            modelBuilder.Entity<Student>()
                .HasOptional(s => s.CurrentInternship).WithRequired(s => s.Intern);

            modelBuilder.Entity<Thesis>()
                .HasOptional(s => s.Application).WithRequired(s => s.Thesis);

            modelBuilder.Entity<Thesis>()
                .HasOptional(s => s.Evaluation).WithRequired(s => s.Thesis);

            modelBuilder.Entity<Thesis>()
                .HasOptional( s => s.ThesisDocument ).WithOptionalPrincipal();

            modelBuilder.Entity<Thesis>()
                .HasMany(x => x.EvaluationCommittee)
                .WithMany(x => x.InEvaluationCommiteeOf)
                .Map(x =>
                        {
                            x.MapLeftKey("ThesisId");
                            x.MapRightKey("TeacherId");
                            x.ToTable("ThesisisEvalutationCommissions");
                        });

            modelBuilder.Entity<Internship>()
                .HasOptional(i => i.InternshipApplication).WithRequired(i => i.Internship);

            modelBuilder.Entity<Internship>()
                .HasOptional(i => i.InternshipEvaluation).WithRequired(i => i.Internship);

            //modelBuilder.Entity<Internship>()
            //    .HasRequired(s => s.InternshipManager).WithMany(s => s.ManagerOfInternship);

            modelBuilder.Entity<ThesisApplication>()
                .HasRequired( s => s.Manager ).WithMany( s => s.ManagerOf );

            modelBuilder.Entity<ThesisApplication>()
                .HasMany(s => s.Consultants).WithMany(s => s.ConsultantOf);

            modelBuilder.Entity<PhdStudent>()
                .HasOptional( z => z.IndividualPlan).WithRequired(x => x.PhdStudent);

            
        }
    }
}
