using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DissProject.Models
{
    public class ThesisApplication
    {
        public int Id { get; set; }

        public virtual Thesis Thesis { get; set; }

        [NotMapped]
        public virtual Student Student
        {
            get
            {
                if (Thesis == null)
                {
                    return null;
                }
                else
                {
                    return Thesis.Student;
                }
            }
        }
        public virtual int Some { get; set; }

        [ForeignKey("Manager")]
        public virtual int ManagerId { get; set; }
        [ForeignKey("ManagerOf")]
        [Display(Name="Ръководител") ]
        public virtual Teacher Manager { get; set; }

        [ForeignKey( "Consultants" ) ]
        public virtual ICollection<int> ConsultantIds { get; set; }

        [InverseProperty("ConsultantOf")]
        [Display( Name="Консултанти")]
        public virtual ICollection< Person > Consultants { get; set; }

        [Display(Name = "Тема на дипломната работа")]
        public String Subject { get; set; }

        [Display(Name = "Анотация")]
        public String Annotation { get; set; }

        [Display(Name = "Цел на дипломната работа")]
        public String Purpose { get; set; }

        [Display(Name = "Задачи, произтичащи от целта")]
        public String Tasks { get; set; }

        [Display( Name = "Ограничаващи/облекчаващи условия" )]
        public String Constraints { get; set; }

        [Display( Name = "Срок за изпълнение" )]
        public DateTime Deadline { get; set; }

        public ThesisApplication()
        {
            this.Consultants = new HashSet<Person>();
            this.ConsultantIds = new HashSet<int>();
        }
    }
}
