using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace DissProject.Site.Areas.Admin.Models
{
    public class ConfirmAccount
    {
        public int UserId { get; set; }

        public int UserName { get; set; }

        public bool Role { get; set; }


    }

    public class UnapproveUser
    {
        public int UserId { get; set; }

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Потребителско име")]
        public string UserName { get; set; }

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [StringLength(50, ErrorMessage = "Името трябва да бъде с поне {2} символа.", MinimumLength = 1)]
        [Display(Name = "Име")]
        public String FirstName { get; set; }


        [StringLength(50, ErrorMessage = "Фамилията трябва да бъде с поне {2} символа.", MinimumLength = 1)]
        [Display(Name = "Фамилия")]
        public String LastName { get; set; }


        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Роля")]
        public string Role { get; set; }


    }


}