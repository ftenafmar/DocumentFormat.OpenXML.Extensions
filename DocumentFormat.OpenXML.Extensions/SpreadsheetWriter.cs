using System;
using System.Xml;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocumentFormat.OpenXml.Extensions
{
    ///<summary>
    ///Provides the base functionality around Spreadsheets
    ///</summary>
    public class SpreadsheetWriter : AbstractWriter
    {
        //Private constructor - static library of functions
        private SpreadsheetWriter()
        {

        }

        ///<summary>
        ///Inserts a new sheet into a SpreadsheetDocument and returns the name of the sheet
        ///</summary>
        public static WorksheetPart InsertWorksheet(SpreadsheetDocument document)
        {
            return InsertWorksheet(document, "");
        }

        ///<summary>
        ///Given a spreadsheet document and sheet name, inserts a new worksheet
        ///</summary>
        ///<param name="doc">The SpreadsheetDocument to insert the worksheet into.</param>
        ///<param name="sheetname">The name of the new sheet.</param>
        ///<returns>The new WorksheetPart or the existing part if a matching sheetname already exists.</returns>
        public static WorksheetPart InsertWorksheet(SpreadsheetDocument doc, string sheetname)
        {
            // Thanks goes to koshinae@codeplex for this method

            // Get sheets where sheetname is the provided text.
            IEnumerable<Sheet> sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetname);

            // If the specified worksheet does not exist, create it.
            if (sheets.Count() == 0)
            {
                // Find out the next id
                uint newId = (uint)(doc.Package.GetRelationships().Count() + 1);
                string rId = "relId" + newId; // do not set this to rId...

                if (string.IsNullOrEmpty(sheetname)) sheetname = string.Format("Sheet{0}", newId);

                int sheetnumber = doc.WorkbookPart.WorksheetParts.Count() + 1;

                // Create the new worksheetpart
                WorksheetPart wsp = doc.WorkbookPart.AddNewPart<WorksheetPart>(rId);

                // Add important stuff :-)
                doc.WorkbookPart.Workbook.Sheets.AppendChild<Sheet>(new Sheet() { Id = rId, SheetId = newId, Name = sheetname });

                wsp.Worksheet = new Worksheet();
                wsp.Worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                wsp.Worksheet.AppendChild<SheetDimension>(new SheetDimension());
                wsp.Worksheet.AppendChild<SheetViews>(new SheetViews()).AppendChild<SheetView>(new SheetView() { WorkbookViewId = 0 });
                wsp.Worksheet.AppendChild<SheetFormatProperties>(new SheetFormatProperties() { DefaultRowHeight = 15 });
                wsp.Worksheet.AppendChild<SheetData>(new SheetData());
                wsp.Worksheet.AppendChild<PageMargins>(new PageMargins() { Left = 0.7, Right = 0.7, Top = 0.75, Bottom = 0.75, Header = 0.3, Footer = 0.3 });

                // Store the relationship of the workbook and the sheet
                doc.Package.CreateRelationship(new Uri("worksheets/sheet" + sheetnumber + ".xml", UriKind.Relative), System.IO.Packaging.TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", rId);
                doc.Package.Flush(); // !!!
                
                // return the new worksheetpart
                return wsp;
            }

            // Return the sheet we found instead
            return (WorksheetPart) doc.WorkbookPart.GetPartById(sheets.First().Id);
        }

        ///<summary>
        ///Deletes a worksheet from the document
        ///</summary>
        ///<param name="doc">The SpreadsheetDocument to remove the worksheet from.</param>
        ///<param name="sheetName">The name of the sheet to remove.</param>
        ///<returns>A boolean indicating whether the worksheet was successfully removed.</returns>
        public static bool RemoveWorksheet(SpreadsheetDocument doc, string sheetName)
        {
            IEnumerable<Sheet> sheets = doc.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName);
            if (sheets.Count() == 0) return false;

            doc.WorkbookPart.DeletePart(sheets.First().Id);
            doc.WorkbookPart.Workbook.Sheets.RemoveChild<Sheet>(sheets.First());
           
            return true;
        }

        ///<summary>
        ///Given text and a SharedStringTablePart, creates or returns a SharedStringItem with the specified text
        ///</summary> 
        public static uint GetSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            uint i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text) return i;
                i = i + Convert.ToUInt32(1);
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Count = Convert.ToUInt32(i + 1);
            shareStringPart.SharedStringTable.UniqueCount = Convert.ToUInt32(i + 1);

            return i;
        }

        ///<summary>
        ///Creates or gets an existing font with the style information provided.
        ///</summary>
        public static UInt32 CreateFont(SpreadsheetStyle style, WorkbookStylesPart styles)
        {
            Font fontMatch = null;
            UInt32 fontIndex = 0;

            //Loop through and see if there is a matching font style
            foreach (var fontElement in styles.Stylesheet.Fonts)
            {
                Font font = (Font) fontElement;

                //If we have a match then use this font
                if (SpreadsheetStyle.CompareFont(font, style))
                {
                    fontMatch = font;
                    break;
                }
                fontIndex += Convert.ToUInt32(1);
            }

            //Add the new font if not found
            if (fontMatch == null)
            {
                Font font = style.ToFont();
                styles.Stylesheet.Fonts.AppendChild<Font>(font); //Font index already set to new count
                styles.Stylesheet.Fonts.Count = fontIndex + Convert.ToUInt32(1);
            }

            return fontIndex;
        }

        ///<summary>
        ///Creates or gets an existing font with the style information provided.
        ///</summary>
        public static UInt32 CreateFill(SpreadsheetStyle style, WorkbookStylesPart styles)
	    {
		    Fill fillMatch = null;
		    UInt32 fillIndex = 0;

		    //Loop through and see if there is a matching font style
		    foreach (var fillElement in styles.Stylesheet.Fills) 
            {
			    Fill fill = (Fill)fillElement;

			    //If we have a match then use this font
			    if (SpreadsheetStyle.CompareFill(fill, style)) 
                {
				    fillMatch = fill;
				    break;
			    }
			    fillIndex += Convert.ToUInt32(1);
		    }

		    //Add the new fill if not found
		    if (fillMatch == null) 
            {
			    Fill fill = style.ToFill();
                styles.Stylesheet.Fills.AppendChild<Fill>(fill);  //Font index already set to new count
			    styles.Stylesheet.Fills.Count.Value = fillIndex + Convert.ToUInt32(1);
		    }

		    return fillIndex;
	    }

        ///<summary>
        ///Creates or gets an existing font with the style information provided.
        ///</summary>
        public static UInt32 CreateBorder(SpreadsheetStyle style, WorkbookStylesPart styles)
	    {
		    Border borderMatch = null;
		    UInt32 borderIndex = 0;

		    //Loop through and see if there is a matching border style
		    foreach (var borderElement in styles.Stylesheet.Borders) 
            {
			    Border border = (Border) borderElement;

			    //If we have a match then use this font
			    if (SpreadsheetStyle.CompareBorder(border, style)) 
                {
				    borderMatch = border;
				    break; 			    
                }
			    borderIndex += Convert.ToUInt32(1);
		    }

		    //Add the new border if not found
		    if (borderMatch == null) 
            {
			    Border border = style.ToBorder();
			    styles.Stylesheet.Borders.AppendChild<Border>(border); //Font index already set to new count
			    styles.Stylesheet.Borders.Count = borderIndex + Convert.ToUInt32(1);
		    }

		    return borderIndex;
	    }


        ///<summary>
        ///Creates or gets an existing number format.
        ///</summary>
        public static UInt32 CreateNumberFormat(SpreadsheetStyle style, WorkbookStylesPart styles)
        {
            NumberingFormat formatMatch = null;
	        UInt32 formatIndex = 0; //starts at 164

	        //Loop through and see if there is a matching border style
	        if (styles.Stylesheet.NumberingFormats != null) 
            {
		        foreach (var formatElement in styles.Stylesheet.NumberingFormats)
                {
                    var format = (NumberingFormat)formatElement;

			        //If we have a match then use this font
			        if (SpreadsheetStyle.CompareNumberFormat(format, style)) 
                    {
				        formatMatch = format;
				        break;
			        }
			        formatIndex += Convert.ToUInt32(1);
		        }
	        }

	        //Add the new number format if not found
	        if (formatMatch == null) {
		        NumberingFormat format = style.ToNumberFormat();

		        format.NumberFormatId = formatIndex + Convert.ToUInt32(164);
		        if (styles.Stylesheet.NumberingFormats == null) styles.Stylesheet.NumberingFormats = new NumberingFormats(); 
		        styles.Stylesheet.NumberingFormats.AppendChild<NumberingFormat>(format);
		        styles.Stylesheet.NumberingFormats.Count = formatIndex + Convert.ToUInt32(1);
	        }

	        return formatIndex + Convert.ToUInt32(164);
        }

        /// <summary>
        /// Saves the spreadsheet and all related document parts.
        /// </summary>
        public static void Save(SpreadsheetDocument spreadsheet)
        {
            //Save all worksheets
            foreach (WorksheetPart worksheetPart in spreadsheet.WorkbookPart.WorksheetParts)
            {
                SetRowSpans(worksheetPart);
                SetWorksheetDimension(worksheetPart);
                worksheetPart.Worksheet.Save();
            }

            //Save the style information
            WorkbookStylesPart styles = SpreadsheetReader.GetWorkbookStyles(spreadsheet);
            styles.Stylesheet.Save();

            //Save the shared string table part
            if (spreadsheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                SharedStringTablePart shareStringPart = spreadsheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                shareStringPart.SharedStringTable.Save();
            }

            //Save the workbook
            spreadsheet.WorkbookPart.Workbook.Save();
        }

        /// <summary>
        /// Updates the row span attribute to the correct value for the cells contained within its.
        /// </summary>
        public static void SetRowSpans(WorksheetPart worksheetPart)
        {
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            //Loop through all rows in the worksheet
            foreach (var row in sheetData.Elements<Row>())
            {
                if (row.Elements<Cell>().Count() > 0)
                {
                    var startCell = row.Elements<Cell>().First();
                    var endCell = row.Elements<Cell>().Last();

                    string startCol = SpreadsheetReader.ColumnFromReference(startCell.CellReference);
                    string endCol = SpreadsheetReader.ColumnFromReference(endCell.CellReference);
                    int startIndex = SpreadsheetReader.GetColumnIndex(startCol);
                    int endIndex = SpreadsheetReader.GetColumnIndex(endCol);

                    ListValue<StringValue> spans = new ListValue<StringValue>();
                    spans.Items.Add(new StringValue(string.Format("{0}:{1}", startIndex, endIndex)));
                    row.Spans = spans;
                }
            }
        }

        /// <summary>
        /// Updates the row span attribute to the correct value for the cells contained within its.
        /// </summary>
        public static SheetDimension SetWorksheetDimension(WorksheetPart worksheetPart)
        {
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            IEnumerable<Row> rows = sheetData.Elements<Row>();

            string firstCol = string.Empty;
            string lastCol = string.Empty;

            //Loop through all rows in the worksheet
            foreach (Row row in rows)
            {
                Cell startCell = row.Elements<Cell>().First();
                Cell endCell = row.Elements<Cell>().Last();

                string startCol = SpreadsheetReader.ColumnFromReference(startCell.CellReference);
                string endCol = SpreadsheetReader.ColumnFromReference(endCell.CellReference);

                if (firstCol == string.Empty || string.Compare(startCol, firstCol, true) < 0) firstCol = startCol;
                if (lastCol == string.Empty || string.Compare(endCol, lastCol, true) > 0) lastCol = endCol;
            }

            //Write out the dimension value
            SheetDimension dimension = worksheetPart.Worksheet.GetFirstChild<SheetDimension>();

            if (rows.Count() == 0)
            {
                dimension.Reference = new StringValue("A1");
            }
            else
            {
                Row firstRow = rows.First();
                Row lastRow = rows.Last();

                if (object.ReferenceEquals(firstRow, lastRow) && firstCol == lastCol)
                {
                    dimension.Reference = new StringValue(string.Format("{0}{1}", firstCol, firstRow.RowIndex));
                }
                else
                {
                    dimension.Reference = new StringValue(string.Format("{0}{1}:{2}{3}", new object[] { firstCol, firstRow.RowIndex, lastCol, lastRow.RowIndex }));
                }
            }

            return dimension;
        }

        /// <summary>
        /// Writes a culture independent numeric string
        /// </summary>
        public static string ToXmlNumeric(object value)
        {
            if (value.GetType() == typeof(short)) return XmlConvert.ToString((short) value);
            if (value.GetType() == typeof(int)) return XmlConvert.ToString((int) value);
            if (value.GetType() == typeof(long)) return XmlConvert.ToString((long) value);
            if (value.GetType() == typeof(float)) return XmlConvert.ToString((float) value);
            if (value.GetType() == typeof(double)) return XmlConvert.ToString((double) value);
            if (value.GetType() == typeof(decimal)) return XmlConvert.ToString((decimal) value);
            if (value.GetType() == typeof(ushort)) return XmlConvert.ToString((ushort) value);
            if (value.GetType() == typeof(uint)) return XmlConvert.ToString((uint) value);
            if (value.GetType() == typeof(ulong)) return XmlConvert.ToString((ulong) value);

            return value.ToString();
        }

        /// <summary>
        /// Writes a boolean value in the correct format
        /// </summary>
        public static string ToXmlBoolean(object value)
        {
            bool result = false;
            bool.TryParse(value.ToString(), out result);
            return (result) ? "1" : "0";
        }
    }
}
