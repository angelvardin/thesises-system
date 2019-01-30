
using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DissProject.Models
{
    [Table("YearWorkPlanApplications")]
    public class YearWorkPlanApplications
    {
        [Key]
        public int Id { get; set; }
       
        public virtual PhdStudent PhdStudent { get; set; }

        [Display(Name = "Година от подготовката")]
        [Required]
        public int PlanYear { get; set; }

        [Display(Name = "Наименование на работите")]
        [Required]
        public string Title { get; set; }

        [Display(Name = "Съдържание на работите")]
        [Required]
        public string Description { get; set; }

        [Display(Name = "Форми на провеждане")]
        [Required]
        public string FormOfConduct { get; set; }

        [Display(Name = "Форми на отчитане")]
        [Required]
        public string FormOfReport { get; set; }

        [Display(Name = "Срок на изпълнение")]
        [Required]
        [DataType(DataType.Date)]
        public DateTime DueDate { get; set; }

        public Teacher Manager { get; set; }
    }
}