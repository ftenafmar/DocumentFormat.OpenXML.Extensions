using System;
using System.Xml;
using System.IO;
using System.IO.Packaging;
using System.Data;
using System.Collections.Generic;
using System.Reflection;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DocumentFormat.OpenXml.Extensions
{
    [TestClass()]
    public class WordprocessingTests
    {
        public TestContext TestContext {get; set;}

        [TestMethod()]
        public void DocumentCreateTest()
        {
            MemoryStream stream = DocumentReader.Create();
            WordprocessingDocument doc = WordprocessingDocument.Open(stream, true);
            MainDocumentPart mainPart = doc.MainDocumentPart;

            Assert.IsTrue(doc != null, "Document was not created");
            Assert.IsTrue(mainPart != null, "Document main part was not created");
        }

        [TestMethod(), DeploymentItem("Templates\\template.docx")]
        public void DocumentPasteTest()
        {
            MemoryStream stream = DocumentReader.Copy(string.Format("{0}\\Templates\\template.docx", Directory.GetCurrentDirectory()));
            WordprocessingDocument doc = WordprocessingDocument.Open(stream, true);
            MainDocumentPart mainPart = doc.MainDocumentPart;

            DocumentWriter writer = new DocumentWriter(mainPart);

            Text name = writer.PasteText("Koos van der Merwe", "NAME");
            Text age = writer.PasteText("53", "AGE");

            //Save to the memory stream, and then to a file
            writer.Save();

            DocumentWriter.StreamToFile(string.Format("{0}\\templatetest.docx", GetOutputFolder()), stream);

            Assert.IsTrue(name != null, "NAME bookmark not set");
            Assert.IsTrue(age != null, "AGE bookmark not set");
        }


        private string GetOutputFolder()
        {
            return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        }

    }
}