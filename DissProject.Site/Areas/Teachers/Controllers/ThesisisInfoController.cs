using DissProject.Models;
using DissProject.Repository;
using DissProject.Site.Areas.Teachers.Models;
using DissProject.Site.Filters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DocumentGeneration;
using System.Web.Routing;

namespace DissProject.Site.Areas.Teachers.Controllers
{
    [InitializeSimpleMembership]
    [IsApproved]
    [Authorize(Roles = "Teacher")]
    public class ThesisisInfoController : Controller
    {

        IUowData _db;

        Dictionary<ThesisSubjectStatus, string> _displayedStatus;
        Dictionary<string, ThesisSubjectStatus> _status;

        public ThesisisInfoController()
        {
            this._db = new UowData();
            _displayedStatus = new Dictionary<ThesisSubjectStatus, string>();
            _displayedStatus.Add(ThesisSubjectStatus.Aproved, "Одобри");
            _displayedStatus.Add(ThesisSubjectStatus.PartiallyApproved, "Одобри със забележка");
            _displayedStatus.Add(ThesisSubjectStatus.Waiting, "Чакащ");
            _displayedStatus.Add(ThesisSubjectStatus.Denied, "Отхвърли");
            _displayedStatus.Add(ThesisSubjectStatus.Invalid, "Невалидна");

            _status = new Dictionary<string, ThesisSubjectStatus>();

            foreach (var status in _displayedStatus)
            {
                _status.Add(status.Value, status.Key);
        }
                        
        }

        //
        // GET: /Teachers/ThesisisInfo/

        public ActionResult Index()
        { 
            return View();
        }
           
        public ActionResult ApproveThesisis()
        {
           var status = new List<object>();
           status.Add(new { Text = _displayedStatus[ThesisSubjectStatus.Aproved], Value = _displayedStatus[ThesisSubjectStatus.Aproved] });
           status.Add(new { Text = _displayedStatus[ThesisSubjectStatus.PartiallyApproved], Value = _displayedStatus[ThesisSubjectStatus.PartiallyApproved] });
           status.Add(new { Text = _displayedStatus[ThesisSubjectStatus.Waiting], Value = _displayedStatus[ThesisSubjectStatus.Waiting] });
           status.Add(new { Text = _displayedStatus[ThesisSubjectStatus.Denied], Value = _displayedStatus[ThesisSubjectStatus.Denied] });
           ViewBag.Status = status.AsQueryable();


           return View();
        }

        public ActionResult NewUserList()
        {
            List<NewThesisis> newThesisis = new List<NewThesisis>();
            try
            {
                UserProfile person = _db.UserProfiles.All()
                              .Where(x => x.UserName == User.Identity.Name)
                              .Select(x => x).SingleOrDefault();
                var thesisis = _db.Thesis.All().ToList();
                foreach (var item in thesisis)
                {
                    if (item.Application.ManagerId == person.UserId)
                    {
                        if (item.SubjectApplicationStatus == ThesisSubjectStatus.Waiting)
                        {
                            newThesisis.Add(new NewThesisis
                            {
                                Name = item.Student.FirstName + " " + item.Student.LastName,
                                ThesisisTitle = item.Application.Subject,
                                SubjectOfStudies = item.Student.SubjectOfStudies,
                                SubjectApplicationStatus = _displayedStatus[ThesisSubjectStatus.Waiting],
                                UserId = item.Student.PersonId,
                            });

                        }  
                    }
                }
                return View(newThesisis.AsQueryable<NewThesisis>());
            }
            catch( Exception ex )
            {
                return RedirectToAction( "Index" );
            }
        }

        [HttpGet]
        public ActionResult AddThesisEvaluation( int studentId )
        {
            var entities = DISSContext.Current.Entities;
            Student student = entities.Students.GetById( studentId );
            if ( student == null )
            {
                return RedirectToAction( "Error" );
            }

            if ( student.CurrentThesis == null )
            {
                return RedirectToAction( "Error" );
            }

            ThesisEvaluation evaluation = student.CurrentThesis.Evaluation;
            ViewBag.IsEditOperation = true;
            if (evaluation == null)
            {
                evaluation = new ThesisEvaluation();
                evaluation.Thesis = student.CurrentThesis;
                ViewBag.IsEditOperation = false;
            }
            
            var grades = new List<SelectListItem>();
            grades.Add( new SelectListItem{ Text = "Отличен(6)", Value = "6" } );
            grades.Add(new SelectListItem { Text = "Мн. Добър(5)", Value = "5" });
            grades.Add(new SelectListItem { Text = "Добър(4)", Value = "4" });
            grades.Add(new SelectListItem { Text = "Среден(3)", Value = "3" });
            grades.Add(new SelectListItem { Text = "Слаб(2)", Value = "2" });
            ViewBag.PossibleGradesSelectList = grades;

            return View( "AddThesisEvaluation", evaluation);
        }

        [HttpGet]
        public ActionResult ThesisEvaluationDetails(int studentId)
        {
            var entities = DISSContext.Current.Entities;
            Student student = entities.Students.GetById(studentId);
            if ( student == null )
            {
                return RedirectToAction("Error");
            }

            if (student.CurrentThesis == null)
            {
                return RedirectToAction("Error");
            }

            if (student.CurrentThesis.Evaluation == null)
            {
                return RedirectToAction("Error");
            }

            return View(student.CurrentThesis.Evaluation);
        }

        public ActionResult ThesisisAssignedToTeacher()
        {
            List<ThesisisAssignedToMe> model = new List<ThesisisAssignedToMe>();
            try
            {
                UserProfile person = _db.UserProfiles.All()
                              .Where(x => x.UserName == User.Identity.Name)
                              .Select(x => x).SingleOrDefault();
                Teacher teacher = _db.Teachers.GetById(person.UserId);
                Person currentPerson = _db.People.GetById(person.UserId);
            foreach (var item in _db.Thesis.All().ToList())
            {
                    if (item.SubjectApplicationStatus == ThesisSubjectStatus.Denied
                        || item.SubjectApplicationStatus == ThesisSubjectStatus.Invalid ||
                        item.SubjectApplicationStatus == ThesisSubjectStatus.Waiting)
                    {
                        continue;
                    }
                bool isChanged = false;
                ThesisisAssignedToMe thesis = new ThesisisAssignedToMe();
                thesis.IsDiplomant = item.DefenseDate;
                    thesis.UserId = item.Student.PersonId;
                    thesis.Name = item.Student.FirstName + " " + item.Student.LastName;
                    thesis.SubjectOfStudies = item.Student.SubjectOfStudies;
                if (item.Application.Consultants.Contains(currentPerson))
                {
                    isChanged = true;
                    thesis.IsConsultant = true;
                }
                if (item.EvaluationCommittee.Contains(teacher))
                {
                    isChanged = true;
                    thesis.IsEvaluator = true;
                }
                if (item.Application.ManagerId == teacher.PersonId)
                {
                     isChanged = true;
                     thesis.IsManager = true;
                }

                if (isChanged == true)
                {
                        model.Add(thesis);
                }


                }

                return View(model.AsQueryable<ThesisisAssignedToMe>());
                }
            catch (Exception)
            {
                return View(model.AsQueryable<ThesisisAssignedToMe>());
            }
        }

        [HttpPost]
        public ActionResult AddThesisEvaluation(ThesisEvaluation evaluation, int studentId)
        {
            var entities = DISSContext.Current.Entities;
            Student student = entities.Students.GetById( studentId );
            if ( student == null )
            {
                return RedirectToAction( "Index" );
            }

            Thesis studentThesis = entities.Students.GetById(studentId).CurrentThesis;
            studentThesis.Evaluation = evaluation;
            evaluation.Thesis = studentThesis;
            evaluation.Evaluator = DISSContext.Current.CurrentPerson;

            try
            {
                entities.Thesis.Update(studentThesis);
                entities.SaveChanges();
            }
            catch ( Exception ex )
            {
                return RedirectToAction("Error");
            }

            return RedirectToAction("Index");
        }

   
        //// GET: /Teachers/ThesisisInfo/Details/5
        public ActionResult Details( int userId )
        {
            var entities = DISSContext.Current.Entities;
            Student student = entities.Students.GetById(userId);
            if (student == null)
            {
                return RedirectToAction("Index");
            }

            Thesis thesis = student.CurrentThesis;
            if (thesis == null)
            {
                return RedirectToAction("Index");
            }

            return View( thesis );
        }

        public FileResult getThesisEvaluationDocument(int studentId)
        {
            var entities = DISSContext.Current.Entities;
            Student student = entities.Students.GetById(studentId);
            if (student == null)
            {
                return null;
            }

            if (student.CurrentThesis == null)
            {
                return null;
            }

            if (student.CurrentThesis.Evaluation == null)
            {
                return null;
            }

            ThesisEvaluationGenerator generator = new ThesisEvaluationGenerator();
            Document evaluation = generator.CreatePackage(student.CurrentThesis.Evaluation);

            return File(evaluation.Data, "application/octet-stream", evaluation.Filename); 
        }

        public ActionResult AddEvaluationCommission(int thesisId)
        {
            EvaluationCommission model = new EvaluationCommission();
            Thesis thesis = _db.Thesis.GetById(thesisId);
            if (thesis == null)
            {
                return RedirectToAction("Index");
            }

            model.ThesisId = thesisId;
            int id = -1;
            if(thesis.CommitteeChairmanId != null)
            {
                model.CommissionChairman = thesis.CommitteeChairmanId.Value;
               
                id = thesis.CommitteeChairmanId.Value;

            }
            model.DefenseDate = (thesis.DefenseDate == null) ? DateTime.Now : thesis.DefenseDate.Value;
            var evaluationCommitte = thesis.EvaluationCommittee.Where(x => x.PersonId != id);
            ViewBag.Commission = new SelectList(evaluationCommitte, "PersonId", "Names");
            ViewBag.Teachers = new SelectList(this._db.Teachers.All(), "PersonId", "Names");

            //var evaluationCommitte = new List<Teacher>();
            //foreach (var item in thesis.EvaluationCommittee.ToList())
            //{
            //    if (item.PersonId == id)
            //    {
            //        continue;
            //    }
            //    evaluationCommitte.Add(item);
            //}
            return View(model);
        }

        [HttpPost]
        public ActionResult AddEvaluationCommission(EvaluationCommission model, FormCollection collection)
        {
            if (ModelState.IsValid)
            {
                Thesis thesis = _db.Thesis.GetById(model.ThesisId);
                try
                {
                   
                    if (thesis == null)
                    {
                        ModelState.AddModelError("", "Възникна грешка. Моля опитайте пак");
                        return View(model);
                        
                    }

                    Teacher chairman = _db.Teachers.GetById(model.CommissionChairman);
                    if (thesis.EvaluationCommittee.Count > 0)
                    {
                        thesis.EvaluationCommittee.Clear();
                    }
                    if (chairman != null)
                    {
                        thesis.EvaluationCommittee.Add(chairman);
                        thesis.CommitteeChairmanId = chairman.PersonId;
                    }
                    else
                    {
                        ModelState.AddModelError("", "Възникна грешка. Моля опитайте пак");
                        return View(model);
                    }
                    string consultantsIDs = collection["CommissionIds"];
                    string[]  splitted =  consultantsIDs.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                    int parsed = -1;
                    foreach (var part in splitted)
                    {
                        if (Int32.TryParse(part, out parsed))
                        {
                            if (model.CommissionChairman == parsed)
                            {
                                continue;
                            }
                            Teacher teacher = _db.Teachers.GetById(parsed);
                            if (teacher != null)
                            {
                                thesis.EvaluationCommittee.Add(teacher);
                            }
                            else
                            {
                                ModelState.AddModelError("", "Възникна грешка. Моля опитайте пак");
                                return View(model);
                            }
                        }
                    }
                    thesis.DefenseDate = model.DefenseDate;
                    _db.Thesis.Update(thesis);
                    _db.SaveChanges();
                    return RedirectToAction("Details", new RouteValueDictionary(
                       new { controller = "ThesisisInfo", action = "Details", userId = thesis.Student.PersonId }));

                }
                catch (Exception)
                {
                    ViewBag.Commission = new SelectList(thesis.EvaluationCommittee, "PersonId", "Names");
                    ViewBag.Teachers = new SelectList(this._db.Teachers.All(), "PersonId", "Names");
                    ModelState.AddModelError("", "Възникна грешка. Моля опитайте пак");
                    return View(model);
                }
            }

            ViewBag.Commission = new SelectList(this._db.Teachers.All(), "PersonId", "Names");
            ViewBag.Teachers = new SelectList(this._db.Teachers.All(), "PersonId", "Names");
            ModelState.AddModelError("", "Възникна грешка. Моля опитайте пак");
            return View(model);
        }
    }
}
