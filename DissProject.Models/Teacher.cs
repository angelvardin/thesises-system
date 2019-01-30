using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DissProject.Models
{
    public class Teacher : Person
    {
        [Required]
        [StringLength(50)]
        public String Title { get; set; }

        [Required]
        public DateTime DateOfApproval { get; set; } // дата на зачисляване

        [InverseProperty( "Manager" )]
        public virtual ICollection<ThesisApplication> ManagerOf { get; set; }

        [InverseProperty("EvaluationCommittee")]
        public virtual ICollection<Thesis> InEvaluationCommiteeOf { get; set; } //студентите на чиито дипломни работи преподавателят е в комисията по оценяване.

        
        public Teacher()
            :base()
        {
            this.ManagerOf = new HashSet<ThesisApplication>();
            this.InEvaluationCommiteeOf = new HashSet<Thesis>();
        }

    }

}
