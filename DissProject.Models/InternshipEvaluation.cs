using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DissProject.Models
{
    public class InternshipEvaluation
    {
        public int Id { get; set; }
        public virtual Internship Internship { get; set; }
        public virtual Teacher teacher { get; set; }

        [NotMapped]
        public virtual Student InternStudent
        {
            get
            {
                return Internship.Intern;
            }
        }

        [Display(Name = "Постигнати резултати")]
        public String Results { get; set; }

        [Display(Name = "Мнение на преподавателя")]
        public String TeacherOpinion { get; set; }

        [Range(2, 6)]
        [Display(Name = "Оценка")]
        public int Grade { get; set; }
    }
}
