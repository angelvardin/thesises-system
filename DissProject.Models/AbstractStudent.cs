using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DissProject.Models
{
    public enum FormOfEducation
    {
        FullTimeStudy,
        PartTimeStudy,
        IndividualFormOfStudy // самостоятелна форма на обучение
    }

    public abstract class AbstractStudent : Person
    {
        [Required]
        [Display(Name = "Специалност")]
        public String SubjectOfStudies { get; set; } //специалност

        [Required]
        [Display(Name = "Форма на обучение" )]
        public FormOfEducation FormOfEducation { get; set; }

        //[ForeignKey("Internship")]
        //public int InternshipId{ get; set; }
        //public int InternshipId { get; set; }
        public Internship Internship { get; set; }
    }
}
