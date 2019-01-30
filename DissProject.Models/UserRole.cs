using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DissProject.Models
{
    public enum UserRole
    {
        Invalid       = -1,
        Administrator = 0,
        Teacher       = 1,
        Student       = 2,
        PhdStudent    = 3
    }

    public class UserRoleUtilities
    {
        public static String userRoleToString(UserRole role)
        {
            switch (role)
            {
                case UserRole.Administrator: return "Administrator";
                case UserRole.PhdStudent: return "PhdStudent";
                case UserRole.Student: return "Student";
                case UserRole.Teacher: return "Teacher";
                default: return "";
            }
        }

        public static UserRole userRoleFromString(String aUserRoleString)
        {
            switch (aUserRoleString)
            {
                case "Administrator": return UserRole.Administrator;
                case "PhdStudent"   : return UserRole.PhdStudent;
                case "Student": return UserRole.Student;
                case "Teacher": return UserRole.Teacher;
                default: return UserRole.Invalid;
            }
        }
    }
}
