using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DissProject.Models
{
    [Table("People")]
    public abstract class Person
    {
        [Key]
        public int PersonId { get; set; }

        public virtual UserProfile User { get; set; }
        public virtual Department Department { get; set; }

        [Required]
        [StringLength(50)]
        public String FirstName { get; set; }

        [StringLength(50)]
        public String SecondName { get; set; }

        [StringLength(50)]
        public String LastName { get; set; }

        [NotMapped]
        public String Names
        {
            get
            {
                return FirstName + " " + LastName;
            }
        }

        [NotMapped]
        public String AllNames
        {
            get
            {
                return FirstName + " " + SecondName + " " + LastName;
            }
        }

        public String PhoneNumber { get; set; }
        public String Address { get; set; }

        [InverseProperty("Consultants")]
        public virtual ICollection< ThesisApplication > ConsultantOf { get; set; }

        [InverseProperty("Evaluator")]
        public ICollection<ThesisEvaluation> EvaluatorOf { get; set; }

        public DateTime? DateOfEarnedTitle { get; set; }

        public Person()
        {
            this.ConsultantOf = new HashSet<ThesisApplication>();
            this.EvaluatorOf = new HashSet<ThesisEvaluation>();
        }
    }
}
