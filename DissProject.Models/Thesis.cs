using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DissProject.Models
{
    public enum ThesisSubjectStatus
    {
        Invalid,
        [Display(Name="Чака Одобрение")]
        Waiting,
        [Display(Name="Неодобрено")]
        Denied,
        [Display(Name = "Одобрено със забележки")]
        PartiallyApproved,
        [Display(Name = "Одобрено")]
        Aproved,
    };

    [Table("Thesises")]
    public class Thesis
    {
        public virtual Student Student { get; set; }

        [Key]
        public int Id { get; set; }

        public virtual ThesisApplication Application { get; set; } //задание
        public virtual ThesisEvaluation Evaluation { get; set; } // рецензия

        [ Display(Name="Статус на заявлението") ]
        public ThesisSubjectStatus SubjectApplicationStatus { get; set; }

        [NotMapped]
        public bool IsApplicationApproved
        {
            get
            {
                return SubjectApplicationStatus == ThesisSubjectStatus.Aproved
                       || SubjectApplicationStatus == ThesisSubjectStatus.PartiallyApproved;
            }
        }

        [ForeignKey("ThesisDocument")]
        public int? ThesisDocumentId { get; set; }
        [Display(Name = "Дипломна Работа")]
        public virtual Document ThesisDocument { get; set; }

        [ForeignKey("ResumeBulgarian") ]
        public int? ResumeBulgarianId { get; set; }
        [Display(Name = "Резюме на Български")]
        public virtual Document ResumeBulgarian { get; set; }

        [ForeignKey("ResumeEnglish") ]
        public int? ResumeEnglishId { get; set; }
        [Display(Name = "Резюме на Английски")]
        public virtual Document ResumeEnglish { get; set; }

        [ForeignKey("SourceCode")]
        public int? SourceCodeId { get; set; }
        [Display(Name="Сорс Код") ]
        public virtual Document SourceCode { get; set; }

        [Range(2, 6)]
        [ Display( Name="Оценка" )]
        public int? Grade { get; set; }

        [ Display(Name="Комисия за оценяване" )]
        public virtual ICollection<Teacher> EvaluationCommittee { get; set; } // комисия 
  
        public int? CommitteeChairmanId { get; set; }

        [Display(Name="Дата за дипломна защита")]
        public DateTime? DefenseDate { get; set; }

        public Thesis()
        {
            this.EvaluationCommittee = new HashSet<Teacher>();
        }
    }
}
