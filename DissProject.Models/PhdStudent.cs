using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DissProject.Models
{
    public enum PhdStudentStatus
    {
        // Това е статуса на докторанта – зачислен, защитил, прекъснал,
        // отчислен без право на защита и отчислен с право на защита.
        StatusApproved,
        StatusDefended,
        StatusExpelled,                       
        StatusDischargedWithoutRightToDefend, 
        StatusDischargedWithRightToDefend     
    }

    public class PhdStudent : AbstractStudent
    {
        public virtual Teacher DirectorOfStudies { get; set; }

        [StringLength(100)]
        public String Protocol { get; set; } // протоколът, с който е зачислен докторанта.

        public PhdStudentStatus Status { get; set; }
        public String Code { get; set; }             // тип varchar с дължина 70, без стойност по подразбиране. Това е кода на докторанта.
        public DateTime DateOfApproval { get; set; } // дата на зачисляване

        public Document WorkSchedule { get; set; } // общ работен план
        public Document IndividualSchedule { get; set; } // индивидуален план
        public Document DistributedWorkShedule { get; set;  } //работен план по години

        public IndividualPlan IndividualPlan { get; set; }
        
        public PhdStudent()
            :base()
        {

        }
    }
}
