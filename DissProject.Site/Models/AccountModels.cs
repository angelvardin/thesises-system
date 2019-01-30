using DissProject.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;
using System.Globalization;
using System.Web.Security;

namespace DissProject.Site.Models
{

    public class ExternalLogin
    {
        public string Provider { get; set; }
        public string ProviderDisplayName { get; set; }
        public string ProviderUserId { get; set; }
    }


    public class RegisterExternalLoginModel
    {
        [Required]
        [Display(Name = "User name")]
        public string UserName { get; set; }

        public string ExternalLoginData { get; set; }
    }

    public class LocalPasswordModel
    {
        [Required(ErrorMessage = "Полето {0} е задължително")]
        [DataType(DataType.Password)]
        [Display(Name = "Сегашна парола")]
        public string OldPassword { get; set; }

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [StringLength(100, ErrorMessage = "Паролата трябва да бъде с поне {2} символа.", MinimumLength = 6)]
        [DataType(DataType.Password)]
        [Display(Name = "Нова парола")]
        public string NewPassword { get; set; }

        [DataType(DataType.Password)]
        [Display(Name = "Потвърди паролата")]
        [Compare("Password", ErrorMessage = "Полето не се съвпада с паролата.")]
        public string ConfirmPassword { get; set; }
    }

    public class LoginModel
    {
        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Потребителско име")]
        public string UserName { get; set; }

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [DataType(DataType.Password)]
        [Display(Name = "Парола")]
        public string Password { get; set; }

        [Display(Name = "Запомни ме?")]
        public bool RememberMe { get; set; }
    }


    public class RegisterModel
    {
        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Потребителско име")]
        public string UserName { get; set; }

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [StringLength(100, ErrorMessage = "Паролата трябва да бъде с поне {2} символа.", MinimumLength = 6)]
        [DataType(DataType.Password)]
        [Display(Name = "Парола")]
        public string Password { get; set; }

        [DataType(DataType.Password)]
        [Display(Name = "Потвърди паролата")]
        [Compare("Password", ErrorMessage = "Полето не се съвпада с паролата.")]
        public string ConfirmPassword { get; set; }

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

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Роля")]
        public string Role { get; set; }

    }

}
