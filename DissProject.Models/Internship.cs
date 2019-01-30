using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace DissProject.Models
{
    public enum InternshipStatus
    {
        [Display(Name="Чака Одобрение")] 
        Applied,
        [Display(Name = "Одобрено")]
        ApprovedApplication,
        [Display(Name = "Оценено")]
        Evaluated,
        [Display(Name = "Оценката е одобрена")]
        ApprovedEvaluation,
    };

    [Table("Internships")]
    public class Internship
    {
        public Internship()
        {
            //this.Intern = new Student();
            //this.InternshipApplication = new InternshipApplication();
            //this.InternshipEvaluation = new InternshipEvaluation();
        }

        [Key]
        [ForeignKey("Intern")] // Id is PK an FK in the same time
        public int Id { get; set; }

        [Display(Name = "Стажант")]
        public virtual Student Intern { get; set; }

        public virtual InternshipApplication InternshipApplication { get; set; }
        public virtual InternshipEvaluation InternshipEvaluation { get; set; }

        [Display(Name = "Статус на заявлението")]
        public InternshipStatus InternshipStatus { get; set; }

        [Range(2,6)]
        public int? Grade { get; set; }

        [ForeignKey("InternshipManager")]
        public int InternshipManagerId { get; set; }

        //[InverseProperty("ManagerOfInternship")]
        [Display(Name = "Консултант")]
        public Person InternshipManager { get; set; } //удобряващ стажа
    }
}
