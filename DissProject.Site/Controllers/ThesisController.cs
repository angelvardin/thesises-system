using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;

using DissProject.Models;
using DissProject.Repository;
using WebMatrix.WebData;
using System.Data.Entity.Validation;
using System.IO;
using System.Web.Helpers;
using DissProject.Site;
using DocumentGeneration;

namespace DissProject.Site.Controllers
{
    [DissProject.Site.Filters.InitializeSimpleMembership]
    //[DissProject.Site.Filters.IsApproved]
    public class ThesisController : Controller
    {
        IUowData databaseEntities;

        public ThesisController()
        {
            this.databaseEntities = new UowData();
        }

        private enum UploadFileType
        {
            Thesis,
            AnnotationBulgarian,
            AnnotationEnglish,
            SourceCode
        }
        //
        // GET: /Thesis/
       
        //[Authorize( Roles = UserRoleUtilities.userRoleToString( UserRole.Student ) )]
        public ActionResult Index()
        {
            DISSContext context = DISSContext.Current;
            var entities = context.Entities;
            if ( context.CurrentRole == UserRole.Student )
            {
                ViewBag.Consultants = new SelectList(this.databaseEntities.Teachers.All(), "PersonId", "Names");

                // TODO: find a way to get off this fucking workaround
                ViewBag.ThesisFile = null;
                if (context.CurrentStudent.CurrentThesis != null)
                {
                    if (context.CurrentStudent.CurrentThesis.ThesisDocumentId.HasValue)
                    {
                        ViewBag.ThesisFile = entities.Documents.GetById(context.CurrentStudent.CurrentThesis.ThesisDocumentId.Value);
                    }

                    ViewBag.AnnotationBulgarian = null;

                    if (context.CurrentStudent.CurrentThesis.ResumeBulgarianId.HasValue)
                    {
                        ViewBag.AnnotationBulgarian = entities.Documents.GetById(context.CurrentStudent.CurrentThesis.ResumeBulgarianId.Value);
                    }

                    ViewBag.AnnotationEnglish = null;

                    if (context.CurrentStudent.CurrentThesis.ResumeEnglishId.HasValue)
                    {
                        ViewBag.AnnotationEnglish = entities.Documents.GetById(context.CurrentStudent.CurrentThesis.ResumeEnglishId.Value);
                    }

                    ViewBag.SourceCode = null;

                    if (context.CurrentStudent.CurrentThesis.SourceCodeId.HasValue)
                    {
                        ViewBag.SourceCode = entities.Documents.GetById(context.CurrentStudent.CurrentThesis.SourceCodeId.Value);
                    }
                }

                return View( context.CurrentStudent.CurrentThesis );
            }

            return RedirectToAction("Index", "Home");
        }

        private JsonResult UploadFileForPerson( int personId, UploadFileType fileType )
        {
            var entities = DISSContext.Current.Entities;
            Document thesisDocument = new Document();
            string errorString = "";

            if (!Utilities.getUploadedDocument(Request, ref thesisDocument, ref errorString))
            {
                return Json(new { error = errorString });
            }

            Student student = entities.Students.GetById(personId);
            if (student == null)
            {
                return Json("No such student");
            }

            Thesis thesis = student.CurrentThesis;
            if (thesis == null || thesis.Application == null)
            {
                errorString = "No valid thesis";
                return Json(new { error = errorString });
            }

            // TODO: uncoment when approving thesis is implemented
            if (!thesis.IsApplicationApproved)
            {
                return Json("Не може да качвате файлове(молбата ви не е одобрена)");
            }

            entities.Documents.Add(thesisDocument);
            entities.SaveChanges();

            switch (fileType)
            {
                case UploadFileType.AnnotationBulgarian:
                    {
                        thesis.ResumeBulgarianId = thesisDocument.Id;
                        break;
                    }
                case UploadFileType.Thesis:
                    {
                        thesis.ThesisDocumentId = thesisDocument.Id;
                        break;
                    }
                case UploadFileType.AnnotationEnglish:
                    {
                        thesis.ResumeEnglishId = thesisDocument.Id;
                        break;
                    }
                case UploadFileType.SourceCode:
                    {
                        thesis.SourceCodeId = thesisDocument.Id;
                        break;
                    }
                 
            }

            thesis.Student = student;
            student.CurrentThesis = thesis;

            try
            {
                DISSContext.Current.Entities.Students.Update(student);
                DISSContext.Current.Entities.SaveChanges();
            }
            catch (Exception e)
            {
                errorString = "Error: File not saved to database";
                return Json(new { error = errorString });
            }

            return Json(new { success = true });
        }

        public JsonResult UploadAnnotationEnglishDocument( int personId )
        {
            return UploadFileForPerson(personId, UploadFileType.AnnotationEnglish);
        }

        public JsonResult UploadAnnotationBulgarianDocument( int personId )
        {
            return UploadFileForPerson(personId, UploadFileType.AnnotationBulgarian );
        }

        public JsonResult UploadThesisDocument( int personId )
        {
            return UploadFileForPerson(personId, UploadFileType.Thesis);
        }

        public JsonResult UploadThesisSourceCode(int personId)
        {
            return UploadFileForPerson(personId, UploadFileType.SourceCode );
        }

        // GET: /Thesis/Create
        public ActionResult Create()
        {
            ViewBag.Teachers = new SelectList(this.databaseEntities.Teachers.All(), "PersonId", "Names");
            var peopleQueryResult = from person in this.databaseEntities.People.All()
                                    where person.PersonId != DISSContext.Current.CurrentStudent.PersonId
                                    select person;

            ViewBag.People = new SelectList( peopleQueryResult , "PersonId", "Names");
            ViewBag.IsEditOperation = false;
            return View();
        }

        [HttpGet]
        public FileResult DownloadDocument( int documentId )
        {
            var entities = DISSContext.Current.Entities;
            Document document = entities.Documents.GetById(documentId);
            return File(document.Data, "application/octet-stream", document.Filename );
        }

        //
        // POST: /Thesis/Create
        [HttpPost]
        public ActionResult Create(ThesisApplication application, FormCollection collection )
        {
            Student student = DISSContext.Current.CurrentStudent;
            var entities = DISSContext.Current.Entities;
            if ( student == null)
            {
                return RedirectToAction("Index");
            }

            try
            {
                string consultantsIDs = collection["ConsultantsIds"];
                string[]  splitted =  consultantsIDs.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                int parsed = -1;
                foreach( string part in splitted )
                {
                    if ( Int32.TryParse(part, out parsed) )
                    {
                        Person person = entities.People.GetById( parsed );
                        if ( person != null )
                        {
                            application.Consultants.Add(person);
                            application.ConsultantIds.Add( parsed );
                        }
                        //application.ConsultantIds.Add(parsed);
                    }
                }

                Thesis thesis = student.CurrentThesis;
                if (thesis == null)
                {
                    thesis = new Thesis();
                    thesis.Application = application;
                    thesis.SubjectApplicationStatus = ThesisSubjectStatus.Waiting;
                    thesis.Student = student;
                    student.CurrentThesis = thesis;
                    entities.Students.Update(student);
                }
                else
                {
                    thesis.Application = application;
                    entities.Thesis.Update(thesis);
                }
                
                entities.SaveChanges();

                return RedirectToAction("Index");
            }
            catch( DbEntityValidationException e )
            {
                return View();
            }
        }

        public ActionResult ShowEditThesisApplication( int studentId )
        {
            var entities = DISSContext.Current.Entities;
            if (DISSContext.Current.CurrentStudent.PersonId != studentId )
            {
                return RedirectToAction( "Error" );
            }

            Student student = DISSContext.Current.CurrentStudent;
            if (student.CurrentThesis == null)
            {
                return RedirectToAction( "Error" );
            }

            if ( student.CurrentThesis.Application == null )
            {
                return RedirectToAction( "Error" );
            }

            if (student.CurrentThesis.IsApplicationApproved)
            {
                return RedirectToAction("Error");
            }

            ViewBag.Teachers = new SelectList(this.databaseEntities.Teachers.All(), "PersonId", "Names");
            var peopleQueryResult = from person in this.databaseEntities.People.All()
                                    where person.PersonId != DISSContext.Current.CurrentStudent.PersonId
                                    select person;

            ViewBag.People = new SelectList(peopleQueryResult, "PersonId", "Names");
            ViewBag.IsEditOperation = true;
            return View("Create", student.CurrentThesis.Application );
        }

        //
        // POST: /Thesis/Edit/5
        [HttpPost]
        public ActionResult EditThesisApplication(ThesisApplication application, FormCollection collection )
        {
            Student student = DISSContext.Current.CurrentStudent;
            var entities = DISSContext.Current.Entities;
            if (student == null)
            {
                return RedirectToAction("Index");
            }
            try
            {
                string consultantsIDs = collection["ConsultantsIds"];
                string[] splitted = consultantsIDs.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                int parsed = -1;
                foreach (string part in splitted)
                {
                    if (Int32.TryParse(part, out parsed))
                    {
                        Person person = entities.People.GetById(parsed);
                        if (person != null)
                        {
                            application.Consultants.Add(person);
                            application.ConsultantIds.Add(parsed);
                        }
                        //application.ConsultantIds.Add(parsed);
                    }
                }

                Thesis thesis = student.CurrentThesis;
                thesis.Application = application;

                entities.Students.Update(student);
                entities.SaveChanges();

                return RedirectToAction("Index");
            }
            catch (DbEntityValidationException e)
            {
                return View();
            }
        }

        [HttpPost]
        public FileResult GetThesisApplicationDocument()
        {
            Student currentStudent = DISSContext.Current.CurrentStudent;
            if (currentStudent == null)
            {
                return null;
            }

            if (currentStudent.CurrentThesis.Application == null)
            {
                return null;
            }

            ThesisApplicationGenerator generator = new ThesisApplicationGenerator();
            Document doc = generator.CreatePackage(currentStudent.CurrentThesis.Application);
            return File( doc.Data, "application/octet-stream", doc.Filename);
        }
    }
}
