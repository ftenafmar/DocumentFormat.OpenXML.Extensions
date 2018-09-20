using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocumentFormat.OpenXml.Extensions
{
    public class DocumentReader: AbstractReader 
    {
        //Private constructor 
        private DocumentReader()
        {
        }

        /// <summary>
        /// Returns a new spreadsheet document as a stream from the blank spreadsheet template.
        ///</summary>
        public static MemoryStream Create()
        {
            return GetEmbeddedResourceStream("Templates\\blank.docx");
        }
    }
}
