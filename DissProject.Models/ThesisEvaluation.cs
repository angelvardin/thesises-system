using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DissProject.Models
{
    public class ThesisEvaluation
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

        public virtual Person Evaluator { get; set; }

        // общи критерии
        [Display( Name="Теоретична Обосновка" )]
        [Range(2,6)]
        public int TheoreticalConsistencyGrade{ get; set;}

        [Display(Name = "Собствени идеи")]
        [Range(2, 6)]
        public int PersonalIdeasGrade { get; set; }

        [Display( Name="Изпълнение на заданието" )]
        [Range(2, 6)]
        public int ExecutionGrade { get; set; }

        [Display(Name = "Стил и оформление")]
        [Range(2, 6)]
        public int StyleGrade { get; set; }

        [NotMapped]
        [Display(Name="Общи критерии")]
        public double CommonCriteriaGrade
        {
            get
            {
                return (TheoreticalConsistencyGrade
                         + PersonalIdeasGrade
                         + ExecutionGrade
                         + StyleGrade) / 4.0;
            }            
        }

        //реализация
        [Display(Name = "Структура и архитектура")]
        [Range(2, 6)]
        public int StructuralGrade { get; set; }

        [Display(Name = "Функционалност")]
        [Range(2, 6)]
        public int FunctionalityGrade { get; set; }

        [Display(Name = "Надеждност")]
        [Range(2, 6)]
        public int ReliabilityGrade { get; set; }

        [Display(Name = "Документация")]
        [Range(2, 6)]
        public int DocumentationGrade { get; set; }

        [NotMapped]
        public double RealizationGrade
        {
            get
            {
                return (StructuralGrade
                         + FunctionalityGrade
                         + ReliabilityGrade
                         + DocumentationGrade) / 4.0;
            }
        }

        // eкспериментална част
        [Display(Name = "Описание на експеримента")]
        [Range(2, 6)]
        public int ExperimentDescriptionGrade { get; set; }

        [Display(Name = "Представяне на резултатите")]
        [Range(2, 6)]
        public int ResultsPresentationGrade { get; set; }

        [Display(Name = "Интерпретация на резултатите")]
        [Range(2, 6)]
        public int InterpretationOfResultsGrade { get; set; }

        [NotMapped]
        public double ExperimentalPartGrade 
        {
            get
            {
                return ( ExperimentDescriptionGrade
                         + ResultsPresentationGrade
                         + InterpretationOfResultsGrade ) / 3.0;
            }
        }

        [NotMapped]
        public double ReccomnendedOVerallGrade
        {
            get
            {
                return ( CommonCriteriaGrade
                         + RealizationGrade
                         + ExperimentalPartGrade) / 3.0;
            }
        }

        [Display(Name="Обща оценка")]
        public double OverallGrade { get; set; }

        [Display(Name="Обобщено мнениее")]
        public string OverallOpinion { get; set; }

        [Display(Name="Questions")]
        public string Questions { get; set; }
    }
}
