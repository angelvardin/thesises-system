using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace DissProject.Models
{
    public class InternshipApplication
    {
        public int Id { get; set; }
        public virtual Internship Internship { get; set; }
        
        [NotMapped]
        public Student Student
        {
            get
            {
                return Internship.Intern;
            }
        }

        [Display(Name = "Тема на стажа")]
        public String InternshipOffer { get; set; }

        [Display(Name = "Анотация")]
        public String Anotation { get; set; }

        [Display(Name = "Цел на стажа")]
        public String Purpose { get; set; }

        [ForeignKey("Consultant")]
        public virtual int ConsultantId { get; set; }
        public virtual Teacher Consultant { get; set; }


    }
}
