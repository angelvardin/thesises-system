using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DissProject.Models;
using System.IO;

namespace DissProject.Site
{
    public static class Utilities
    {
        public static string InternshipStatusString( InternshipStatus i )
        {
            switch (i)
            {
                case InternshipStatus.Applied: return "Чака Одобрение";
                case InternshipStatus.ApprovedApplication: return "Одобрено";
                case InternshipStatus.Evaluated: return "Оценено";
                case InternshipStatus.ApprovedEvaluation: return "Оценката е одобрена";
                default: return "";
            }
        }

        public static string ThesisSubjectStatusString( ThesisSubjectStatus a )
        {
            switch (a)
            {
                case ThesisSubjectStatus.Waiting: return "Чака Одобрение";
                case ThesisSubjectStatus.Denied: return "Отказан";
                case ThesisSubjectStatus.Aproved: return "Одобрен";
                case ThesisSubjectStatus.PartiallyApproved: return "Одобрен със забележки";
                default: return "";
            }
        }

        public static bool getUploadedDocument( HttpRequestBase request, ref Document document, ref string errorString )
        {
            Stream fileStream = null;
            byte[] documentBytes;
            String fullFileName;

            if (String.IsNullOrEmpty(request["qqfile"]))
            {
                //This works with IE
                HttpPostedFileBase httpPostedFileBase = request.Files[0] as HttpPostedFileBase;
                fileStream = httpPostedFileBase.InputStream;
                fullFileName = httpPostedFileBase.FileName;
            }
            else
            {
                // other browsers
                fileStream = request.InputStream;
                fullFileName = request["qqfile"];
            }

            if ( fileStream == null )
            {
                errorString = "Couldn't get filestream from request";
                return false;
            }

            using (var reader = new BinaryReader(fileStream))
            {
                documentBytes = reader.ReadBytes((int)fileStream.Length);
            }

            document.Data = documentBytes;
            document.DateCreated = DateTime.Now;
            document.Filename = fullFileName;
            document.DateLastModified = DateTime.Now;
            return true;
        }
    }
}