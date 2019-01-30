using DissProject.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace DissProject.Site.Areas.Teachers.Models
{
    public class ThesisDetailsModel
    {
        [Display(Name = "Тема на дипломната работа")]
        public String Subject { get; set; }

        [Display(Name = "Анотация")]
        public String Annotation { get; set; }

        [Display(Name = "Цел на дипломната работа")]
        public String Purpose { get; set; }

        [Display(Name = "Задачи, произтичащи от целта")]
        public String Tasks { get; set; }

        [Display(Name = "Ограничаващи/облекчаващи условия")]
        public String Constraints { get; set; }

        [Display(Name = "Срок за изпълнение")]
        public DateTime Deadline { get; set; }
    }

    public class NewThesisis
    {
        public int UserId { get; set; }


        [Required(ErrorMessage = "Полето {0} е задължително")]
        [StringLength(50, ErrorMessage = "Името трябва да бъде с поне {2} символа.", MinimumLength = 1)]
        [Display(Name = "Име")]
        public String Name { get; set; }


        [StringLength(50, ErrorMessage = "Фамилията трябва да бъде с поне {2} символа.", MinimumLength = 1)]
        [Display(Name = "Фамилия")]
        public String ThesisisTitle { get; set; }

        [Display(Name = "Статус")]
        public object SubjectApplicationStatus { get; set; }


        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Специалност")]
        public string SubjectOfStudies { get; set; }
    }

    public class ThesisisAssignedToMe
    {
        public int UserId { get; set; }

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [StringLength(50, ErrorMessage = "Името трябва да бъде с поне {2} символа.", MinimumLength = 1)]
        [Display(Name = "Име")]
        public String Name { get; set; }


        [StringLength(50, ErrorMessage = "Фамилията трябва да бъде с поне {2} символа.", MinimumLength = 1)]
        [Display(Name = "Фамилия")]
        public String ThesisisTitle { get; set; }

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Специалност")]
        public string SubjectOfStudies { get; set; }

        public bool IsManager { get; set; }

        /// <summary>
        /// Рецензент
        /// </summary>
        public bool IsEvaluator { get; set; }

        public bool IsConsultant { get; set; }

        public Nullable<DateTime> IsDiplomant { get; set; }
    }


    public class EvaluationCommission
    {
        public int ThesisId { get; set; }

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Председател на комисията")]
        public int CommissionChairman { get; set; }

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Дата на защита")]
        public DateTime DefenseDate { get; set; }
    }
}