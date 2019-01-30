using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DissProject.Models
{
    public class Student : AbstractStudent
    {
        public String InternshipRating { get; set; } // evaluation of the internship the student has taken

        public virtual Thesis CurrentThesis { get; set; }

        public virtual Internship CurrentInternship { get; set; }
      
        [Required]
        public int FacultyNumber { get; set; }

        [Required]
        public int GraduationYear { get; set; }

        public String WorkCompany { get; set; }

        
        public Student()
            :base()
        {

        }
    }
}
