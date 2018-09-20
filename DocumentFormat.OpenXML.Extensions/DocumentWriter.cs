using System;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.IO;
using System.IO.Packaging;
using System.Reflection;
using System.Collections.Generic;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentFormat.OpenXml.Extensions
{
    ///<summary>
    ///Provides the base functionality around WordprocessingDocuments as a set of static functions
    ///</summary>
    public class DocumentWriter : AbstractWriter
    {
        private MainDocumentPart _documentPart;

        public DocumentWriter(MainDocumentPart document)
        {
            if (document == null) throw new ArgumentNullException("document");
            _documentPart = document;
        }

        /// <summary>
        /// Pastes the text provided at a bookmark in the document body.
        /// </summary>
        public Text PasteText(string text, string bookmarkName)
        {
            return PasteText(text, bookmarkName, _documentPart);
        }

        /// <summary>
        /// Pastes a dictionary containing bookmark-text key pair values into the document.
        /// </summary>
        /// <remarks>Bookmarks that are found in the document are removed from the dictionary.</remarks>
        public void PasteText(Dictionary<string, string> bookmarkValues)
        {
            PasteText(bookmarkValues, _documentPart);
        }

        /// <summary>
        /// Saves all elements in the main document part
        /// </summary>
        public void Save()
        {
            _documentPart.Document.Save();
        }

        /// <summary>
        /// Pastes the text provided at a bookmark in the document body.
        /// </summary>
        public static Text PasteText(string text, string bookmarkName, MainDocumentPart documentPart)
        {
            Body body = documentPart.Document.GetFirstChild<Body>();
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();

            //Get all paragraphs of text
            foreach (Paragraph para in paras)
            {
                IEnumerable<BookmarkStart> bookMarkStarts = para.Elements<BookmarkStart>();
                IEnumerable<BookmarkEnd> bookMarkEnds = para.Elements<BookmarkEnd>();

                //Get the id of the bookmark start to find the bookmark end
                foreach (BookmarkStart bookMarkStart in bookMarkStarts)
                {
                    if (bookMarkStart.Name == bookmarkName)
                    {
                        string id = bookMarkStart.Id.Value;
                        BookmarkEnd bookmarkEnd = bookMarkEnds.Where(i => i.Id.Value == id).First();

                        //Create a new run and text
                        var textElement = new Text(text);
                        var runElement = new Run(textElement);

                        para.InsertAfter(runElement, bookmarkEnd);

                        return textElement;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Pastes a dictionary containing bookmark-text key pair values into a document.
        /// </summary>
        /// <remarks>Bookmarks that are found in the document are removed from the dictionary.</remarks>
        public static void PasteText(Dictionary<string, string> bookmarkValues, MainDocumentPart documentPart)
        {
            if (bookmarkValues == null) return;

            Body body = documentPart.Document.GetFirstChild<Body>();
            IEnumerable<Paragraph> paras = body.Elements<Paragraph>();

            //Get all paragraphs of text
            foreach (Paragraph para in paras)
            {
                IEnumerable<BookmarkStart> bookMarkStarts = para.Elements<BookmarkStart>();
                IEnumerable<BookmarkEnd> bookMarkEnds = para.Elements<BookmarkEnd>();

                //Get the id of the bookmark start to find the bookmark end
                foreach (BookmarkStart bookMarkStart in bookMarkStarts)
                {
                    if (bookmarkValues.ContainsKey(bookMarkStart.Name))
                    {
                        string id = bookMarkStart.Id.Value;
                        BookmarkEnd bookmarkEnd = bookMarkEnds.Where(i => i.Id.Value == id).First();

                        //Create a new run and text
                        var textElement = new Text(bookmarkValues[bookMarkStart.Name]);
                        var runElement = new Run(textElement);

                        para.InsertAfter(runElement, bookmarkEnd);
                        bookmarkValues.Remove(bookMarkStart.Name);
                    }
                }
            }
        }

        /// <summary>
        /// Saves all elements in the main document part
        /// </summary>
        public static void Save(MainDocumentPart documentPart)
        {
            documentPart.Document.Save();
        }
    }

}