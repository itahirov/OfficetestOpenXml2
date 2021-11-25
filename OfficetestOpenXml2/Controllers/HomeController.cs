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
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";
            /*DocumentFormat.OpenXml.Packaging.WordprocessingDocument wordprocessingDocument = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open("test.docx", true);
            DocumentFormat.OpenXml.Wordprocessing.Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
            // Add new text.
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text("nigigase"));*/
            //string strDoc = @"C:\Users\v-istahi\Documents\temp1.docx";
            string strDoc = @"C:\inetpub\wwwroot\temp1.docx";
            //string strDoc = @"D:\home\wwwroot\temp1.docx";
            string strTxt = "Append text in body - OpenAndAddTextToWordDocument";
            DateTime dt = DateTime.Now;
            strTxt = dt.ToString("yyyMMMddhhmm");
            OpenAndAddTextToWordDocument(strDoc, strTxt);

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}