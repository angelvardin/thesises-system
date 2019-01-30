using DissProject.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace DissProject.Site.Models
{
    public class StudentInfo
    {

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Факултетен номер")]
        public int FacultyNumber { get; set; }

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Година")]
        public int GraduationYear { get; set; }

        /// <summary>
        /// Специалност
        /// </summary>
        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Специалност")]
        public String SubjectOfStudies { get; set; }

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Форма на обучение")]
        public FormOfEducation FormOfEducation { get; set; }
    }

    public class TeacherInfo
    {
        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Титла")]
        public string TeacherTitle { get; set; }

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Катедра")]
        public int Department { get; set; }

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Дата на зачисляване")]
        [DataType(DataType.DateTime, ErrorMessage = "Полето {0} трябва да е дата")]
        public DateTime DateOfApproval { get; set; }
    }

    public class PhdStudentInfo
    {
        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Протокол")]
        [StringLength(40, ErrorMessage = "Максималната дължина е {0} символа")]
        public string Protocol { get; set; }

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Код")]
        [StringLength(70, ErrorMessage = "Максималната дължина е {0} символа")]
        public string Code { get; set; }

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Дата на зачисляване")]
        [DataType(DataType.DateTime, ErrorMessage = "Полето {0} трябва да е дата")]
        public DateTime DateOfApproval { get; set; }

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Специалност")]
        public String SubjectOfStudies { get; set; }

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Форма на обучение")]
        public FormOfEducation FormOfEducation { get; set; }

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Катедра")]
        public int Department { get; set; }
    }

    public class PersonInfo
    {
        [Required(ErrorMessage = "Полето {0} е задължително")]
        [StringLength(50, ErrorMessage = "Името трябва да бъде с поне {2} символа.", MinimumLength = 1)]
        [Display(Name = "Име")]
        public String FirstName { get; set; }


        [StringLength(50, ErrorMessage = "Презимето трябва да бъде с поне {2} символа.", MinimumLength = 1)]
        [Display(Name = "Презиме")]
        public String SecondName { get; set; }

        [StringLength(50, ErrorMessage = "Фамилията трябва да бъде с поне {2} символа.", MinimumLength = 1)]
        [Display(Name = "Фамилия")]
        public String LastName { get; set; }

        [Display(Name = "Телефонен номер")]
        [Phone]
        public String PhoneNumber { get; set; }

        [Display(Name = "Адрес")]
        [StringLength(50, ErrorMessage = "Адресът трябва да бъде с поне {2} символа.", MinimumLength = 3)]
        public String Address { get; set; }
    }
    public class EditStudentInfo : StudentInfo
    {
        [Required(ErrorMessage = "Полето {0} е задължително")]
        [StringLength(50, ErrorMessage = "Името трябва да бъде с поне {2} символа.", MinimumLength = 1)]
        [Display(Name = "Име")]
        public String FirstName { get; set; }


        [StringLength(50, ErrorMessage = "Презимето трябва да бъде с поне {2} символа.", MinimumLength = 1)]
        [Display(Name = "Презиме")]
        public String SecondName { get; set; }

        [StringLength(50, ErrorMessage = "Фамилията трябва да бъде с поне {2} символа.", MinimumLength = 1)]
        [Display(Name = "Фамилия")]
        public String LastName { get; set; }

        [Display(Name = "Телефонен номер")]
        [Phone]
        public String PhoneNumber { get; set; }

        [Display(Name = "Адрес")]
        [StringLength(50, ErrorMessage = "Адресът трябва да бъде с поне {2} символа.", MinimumLength = 3)]
        public String Address { get; set; }
    }

    public class EditPhdStudentInfo : PhdStudentInfo
    {
        [Required(ErrorMessage = "Полето {0} е задължително")]
        [StringLength(50, ErrorMessage = "Името трябва да бъде с поне {2} символа.", MinimumLength = 1)]
        [Display(Name = "Име")]
        public String FirstName { get; set; }


        [StringLength(50, ErrorMessage = "Презимето трябва да бъде с поне {2} символа.", MinimumLength = 1)]
        [Display(Name = "Презиме")]
        public String SecondName { get; set; }

        [StringLength(50, ErrorMessage = "Фамилията трябва да бъде с поне {2} символа.", MinimumLength = 1)]
        [Display(Name = "Фамилия")]
        public String LastName { get; set; }

        [Display(Name = "Телефонен номер")]
        [Phone]
        public String PhoneNumber { get; set; }

        [Display(Name = "Адрес")]
        [StringLength(50, ErrorMessage = "Адресът трябва да бъде с поне {2} символа.", MinimumLength = 3)]
        public String Address { get; set; }
    }

    public class EditTeacherInfo : TeacherInfo
    {
        [Required(ErrorMessage = "Полето {0} е задължително")]
        [StringLength(50, ErrorMessage = "Името трябва да бъде с поне {2} символа.", MinimumLength = 1)]
        [Display(Name = "Име")]
        public String FirstName { get; set; }


        [StringLength(50, ErrorMessage = "Презимето трябва да бъде с поне {2} символа.", MinimumLength = 1)]
        [Display(Name = "Презиме")]
        public String SecondName { get; set; }

        [StringLength(50, ErrorMessage = "Фамилията трябва да бъде с поне {2} символа.", MinimumLength = 1)]
        [Display(Name = "Фамилия")]
        public String LastName { get; set; }

        [Display(Name = "Телефонен номер")]
        [Phone]
        public String PhoneNumber { get; set; }

        [Display(Name = "Адрес")]
        [StringLength(50, ErrorMessage = "Адресът трябва да бъде с поне {2} символа.", MinimumLength = 3)]
        public String Address { get; set; }
    }


}