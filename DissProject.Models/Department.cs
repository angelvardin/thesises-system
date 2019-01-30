using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DissProject.Models
{
    [Table("Departments")]
    public class Department
    {
        public int Id { get; set; }
        public String Description { get; set; }

        public virtual ICollection<Person> People { get; set; }

        public Department()
        {
            this.People = new HashSet<Person>();
        }
    }
}
