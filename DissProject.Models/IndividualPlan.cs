using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DissProject.Models
{
    [Table("IndividualPlan")]
    public class IndividualPlan
    {
        [Key]
        public int Id { get; set; }

        public virtual PhdStudent PhdStudent { get; set; }

        [Display(Name = "Срок на завършване на докторантурата")]
        [Required]
        [DataType(DataType.Date), DisplayFormat( DataFormatString="{0:dd/MM/yyyy}", ApplyFormatInEditMode=true )]
        public DateTime GratuationDate { get; set; }

        [Display(Name = "Научна специалност")]
        [Required]
        [StringLength(4000)] 
        public string Specialty { get; set; }

        [Display(Name = "Тема на дисертационната работа")]
        [Required]
        [StringLength(4000)] 
        public string PhdThesisTitle { get; set; }

        [Display(Name = "Индивидуален план за работа на докторанта от Kaтедрения съвет в заседание от")]
        [Required]
        [StringLength(4000)] 
        public string FacultyProtocol { get; set; }

        public Teacher Manager { get; set; }
    }
}
