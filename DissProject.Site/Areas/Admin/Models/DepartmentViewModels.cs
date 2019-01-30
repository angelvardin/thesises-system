using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace DissProject.Site.Areas.Admin.Models
{
    public class DepartmentViewModel
    {
        public int DepartmentId { get; set; }

        [Required(ErrorMessage = "Полето {0} е задължително")]
        [Display(Name = "Име")]
        public string Description { get; set; }
    }
}