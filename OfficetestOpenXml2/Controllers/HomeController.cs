using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficetestOpenXml2.Controllers
{
    public class HomeController : Controller
    {
        public static void OpenAndAddTextToWordDocument(string filepath, string txt)
        {
            // Open a WordprocessingDocument for editing using the filepath.
            WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(filepath, true);

            // Assign a reference to the existing document body.
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

            // Add new text.
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(txt));

            // Close the handle explicitly.
            wordprocessingDocument.Close();
        }

        public static void CreateWordprocessingDocument(string filepath)
        {
            // Create a document by supplying the filepath. 
            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("Create text in body - CreateWordprocessingDocument"));
            }
        }
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";
            string strDoc = @"C:\\inetpub\\wwwroot\\reportword\\temp1.docx";
            string strTxt = "Append text in body - OpenAndAddTextToWordDocument";
            DateTime dt = DateTime.Now;
            strTxt = dt.ToString("yyyMMMddhhmm");
            OpenAndAddTextToWordDocument(strDoc, strTxt);

            return View();
        }

        public ActionResult Contact()
        {
            string strDoc = @"C:\\inetpub\\wwwroot\\reportword\\MyDoc.docx\\temp1.docx";
            //string strDoc = @"temp1.docx";
            CreateWordprocessingDocument(strDoc);
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}