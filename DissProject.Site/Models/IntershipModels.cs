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
    public class IntershipModels 
    {
        public int Id { get; set; }

      
        public virtual Student Student { get; set; }
        public virtual Teacher Manager { get; set; }
       

        [Display(Name = "Тема на стажа")]
        public String InternshipOffer { get; set; }

        [Display(Name = "Анотация")]
        public String Evaluation { get; set; }
}
}
