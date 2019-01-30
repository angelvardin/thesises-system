using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DissProject.Models
{
    public static class ModelUtilities
    {
        public static string GradeDescription(int aGrade)
        {
            switch (aGrade)
            {
                case 2: return "слаб";
                case 3: return "среден";
                case 4: return "добър";
                case 5: return "мн. добър";
                case 6: return "отличен";
                default: return "";
            }
        }

    }
}
