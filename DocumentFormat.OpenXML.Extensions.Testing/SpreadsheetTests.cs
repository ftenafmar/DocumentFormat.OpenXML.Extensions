using System;
using System.Xml;
using System.IO;
using System.IO.Packaging;
using System.Data;
using System.Collections.Generic;
using System.Reflection;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DocumentFormat.OpenXml.Extensions
{
    [TestClass()]
    public class SpreadsheetTests
    {
        public TestContext TestContext { get; set; }

        [TestMethod()]
        public void InsertSheetTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
         
            var worksheetPart =  SpreadsheetWriter.InsertWorksheet(doc);

            Assert.IsNotNull(worksheetPart, "A worksheet part was not returned.");

            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);
            writer.PasteText("B2", "Hello World");

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\createsheet.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod()]
        public void RemoveSheetTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);

            var result = SpreadsheetWriter.RemoveWorksheet(doc,"Sheet2");

            Assert.IsTrue(result, "A worksheet was not removed from the document.");

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\removesheet.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod()]
        public void WorksheetPasteTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");

            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);

            writer.PasteNumber("D3", "2");
            writer.PasteNumber("D4", "3");
            writer.PasteNumber("D5", "4");

            //Add total without a calc chain
            writer.FindCell("D6").CellFormula = new CellFormula("SUM(D3:D5)");

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\output.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod()]
        public void WorksheetPasteDate()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");

            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);

            Cell cell = writer.PasteDate("D3", new DateTime(2009,12,20,15,23,05));

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\date.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod()]
        public void WorksheetSharedTextTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");

            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);

            writer.PasteSharedText("B2", "Shared Text");
            writer.PasteSharedText("B3", "Shared Text");

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\sharedtext.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod(), DeploymentItem("Templates\\template.xlsx")]
        public void WorksheetCopyTest()
        {
            MemoryStream stream = SpreadsheetReader.Copy(string.Format("{0}\\Templates\\template.xlsx", Directory.GetCurrentDirectory()));
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");

            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);

            writer.PasteNumber("B3", "10");
            writer.PasteNumber("B4", "20");
            writer.PasteNumber("B5", "40");

            //Add total without a calc chain
            writer.FindCell("B6").CellFormula = new CellFormula("SUM(B3:B5)");

            //Change the print area from A1:I30
            writer.SetPrintArea("Sheet1", "A1", "D10");

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\templatetest.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod()]
        public void SpreadsheetColumnIncrementTest()
        {
            Assert.IsTrue(SpreadsheetReader.GetColumnName("A", 1) == "B", "A + 1 failed.");
            Assert.IsTrue(SpreadsheetReader.GetColumnName("B", 2) == "D", "B + 2 failed.");
            Assert.IsTrue(SpreadsheetReader.GetColumnName("Y", 1) == "Z", "Y + 1 failed.");
            Assert.IsTrue(SpreadsheetReader.GetColumnName("Z", 1) == "AA", "Z + 1 failed.");
            Assert.IsTrue(SpreadsheetReader.GetColumnName("AA", 1) == "AB", "AA + 1 failed.");
            Assert.IsTrue(SpreadsheetReader.GetColumnName("AA", 2) == "AC", "AA + 2 failed.");
            Assert.IsTrue(SpreadsheetReader.GetColumnName("AA", 26) == "BA", "AA + 26 failed.");
            Assert.IsTrue(SpreadsheetReader.GetColumnName("AAA", 1) == "AAB", "AAA + 1 failed.");
            Assert.IsTrue(SpreadsheetReader.GetColumnName("AAZ", 1) == "ABA", "AAZ + 1 failed.");
            Assert.IsTrue(SpreadsheetReader.GetColumnName("AZZ", 1) == "BAA", "AZZ + 1 failed.");

            Assert.IsTrue(SpreadsheetReader.GetColumnName("B", -1) == "A", "B - 1 failed.");
            Assert.IsTrue(SpreadsheetReader.GetColumnName("AA", -1) == "Z", "AA - 1 failed.");
            Assert.IsTrue(SpreadsheetReader.GetColumnName("AAA", -1) == "ZZ", "AAA - 1 failed.");
        }

        [TestMethod()]
        public void SpreadsheetReferenceTests()
        {
            Assert.IsTrue(SpreadsheetReader.ReferenceFromRange("A1") == "A1");
            Assert.IsTrue(SpreadsheetReader.ReferenceFromRange("A1:B1") == "A1");
            Assert.IsTrue(SpreadsheetReader.ReferenceFromRange("$A$1:B1") == "$A$1");

            Assert.IsTrue(SpreadsheetReader.ColumnFromReference("$A$1") == "A");
            Assert.IsTrue(SpreadsheetReader.RowFromReference("$A$1").Value == new UInt32Value(Convert.ToUInt32(1)).Value);
        }

        [TestMethod()]
        public void WorksheetDataTablePasteTest()
        {
            DataTable dt = new DataTable();

            //Manually create table for mockup purposes
            dt.TableName = "tblParameters";
            dt.Columns.Add("intID", typeof(int));
            dt.Columns.Add("vchrDescription", typeof(string));
            dt.Columns.Add("dteValidFrom", typeof(DateTime));
            dt.Columns.Add("decPrice", typeof(float));
            dt.Columns.Add("bitFlag", typeof(Boolean));
            dt.AcceptChanges();

            dt.Rows.Add(new object[] { 1, "Parts", new DateTime(1974,1,2), 1.00F, true});
            dt.Rows.Add(new object[] { 2, "Cheque", new DateTime(1974,2,2), 1.50F, false });
            dt.Rows.Add(new object[] { 3, "Products",  new DateTime(1974,3,2), 1.45F, true });
            dt.Rows.Add(new object[] { 4, "Gifts",  new DateTime(1974,4,2), 0.00F, false });
            dt.Rows.Add(new object[] { 5, "DealerRepair", new DateTime(1974,5,2), 2.50F, true });

            //Write to the spreadsheet
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");
            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);

            SpreadsheetStyle style = SpreadsheetReader.GetDefaultStyle(doc);
            style.IsBold = true;
            writer.PasteDataTable(dt, "B3", style);

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\datatable.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod()]
        public void WorksheetDataTablePasteColumnsTest()
        {
            DataTable dt = new DataTable();

            //Manually create table for mockup purposes
            dt.TableName = "tblParameters";
            dt.Columns.Add("intID", typeof(int));
            dt.Columns.Add("vchrDescription", typeof(string));
            dt.Columns.Add("decPrice", typeof(float));
            dt.AcceptChanges();

            dt.Rows.Add(new object[] { 1, "Parts", 1.00F});
            dt.Rows.Add(new object[] { 2, "Cheque",  1.50F});
            dt.Rows.Add(new object[] { 3, "Products", 1.45F });
            dt.Rows.Add(new object[] { 4, "Gifts", 0.00F });
            dt.Rows.Add(new object[] { 5, "DealerRepair", 2.50F });

            //Write to the spreadsheet
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");
            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);

            writer.PasteDataTable(dt, "B3", new List<string>(new string[] { "vchrDescription" }));

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\datatablepartial.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod(), DeploymentItem("Templates\\template.xlsx")]
        public void WorksheetDataTablePasteColumnsTemplate()
        {
            DataTable dt = new DataTable();

            //Manually create table for mockup purposes
            dt.TableName = "tblParameters";
            dt.Columns.Add("intID", typeof(int));
            dt.Columns.Add("vchrDescription", typeof(string));
            dt.Columns.Add("decPrice", typeof(float));
            dt.AcceptChanges();

            dt.Rows.Add(new object[] { 1, "Parts", 1.00F });
            dt.Rows.Add(new object[] { 2, "Cheque", 1.50F });
            dt.Rows.Add(new object[] { 3, "Products", 1.45F });
            dt.Rows.Add(new object[] { 4, "Gifts", 0.00F });
            dt.Rows.Add(new object[] { 5, "DealerRepair", 2.50F });

            //Write to the spreadsheet
            MemoryStream stream = SpreadsheetReader.Copy(string.Format("{0}\\Templates\\template.xlsx", Directory.GetCurrentDirectory()));
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");
            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);

            writer.PasteDataTable(dt, "B3", new List<string>(new string[] { "vchrDescription" }));

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\datatabletemplate.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod()]
        public void WorksheetPasteValuesTest()
        {
            List<string> values = new List<string>(new string[] { "alpha", "beta", "charlie", "delta" });

            //Write to the spreadsheet
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");

            WorksheetWriter.PasteValues(doc, worksheetPart, "A", 2, values, CellValues.String, null);

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\pastevalues.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod()]
        public void SpreadsheetGetDefaultStyleTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            SpreadsheetStyle style = SpreadsheetReader.GetDefaultStyle(doc);

            Assert.IsTrue(style != null, "Default style not found.");
        }

        [TestMethod()]
        public void WorksheetCreateFontTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);

            WorkbookStylesPart styles = SpreadsheetReader.GetWorkbookStyles(doc);

            SpreadsheetStyle defaultStyle = SpreadsheetReader.GetDefaultStyle(doc);
            uint index = SpreadsheetWriter.CreateFont(defaultStyle, styles);
        }

        [TestMethod()]
        public void WorksheetCreateFillTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);

            WorkbookStylesPart styles = SpreadsheetReader.GetWorkbookStyles(doc);

            SpreadsheetStyle defaultStyle = SpreadsheetReader.GetDefaultStyle(doc);
            uint index = SpreadsheetWriter.CreateFill(defaultStyle, styles);
        }

        [TestMethod()]
        public void WorksheetAddStyleTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");

            SpreadsheetStyle defaultStyle = SpreadsheetReader.GetDefaultStyle(doc);
            defaultStyle.IsItalic = true;
            defaultStyle.IsBold = true;
            defaultStyle.IsUnderline = true;
            defaultStyle.SetColor("FF0000");
            //r = 255, g = 0, b = 0 (Red)
            defaultStyle.SetBackgroundColor("DDDDDD");
            //(light grey)
            defaultStyle.SetBorder("00FF00", BorderStyleValues.Medium);
            //(green medium border)

            WorksheetWriter.PasteText(doc, worksheetPart, "E", 5, "Hello world");
            WorksheetWriter.SetStyle(defaultStyle, doc, worksheetPart, "E", 5);

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\styled.xlsx", GetOutputFolder()), stream);
        }
        
        [TestMethod]
        public void WorksheetAddAlignmentTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");
            SpreadsheetStyle style = SpreadsheetReader.GetDefaultStyle(doc);
            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);

            style.IsWrapped = true;
            writer.PasteText("E5", "Wrapped text", style);

            style.IsWrapped = false;
            style.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            writer.PasteText("E7", "Aligned Test", style);

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\wrapped.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod()]
        public void WorksheetNumberFormatTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");
            SpreadsheetStyle style = SpreadsheetReader.GetDefaultStyle(doc);
            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);

            style.FormatCode = "0.00";
            writer.PasteNumber("B3", "123", style);

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\numberformat.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod()]
        public void WorksheetFindCellTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");

            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);

            //Test that the findcell function can accurately add cells past z
            writer.PasteText("A1","A1");
            writer.PasteText("Z1","Z1");
            writer.PasteText("Z2","Z2");
            writer.PasteText("AA1","AA1");
            writer.PasteText("AA2","AA2");
            writer.PasteText("BA1","BA1");
            writer.PasteText("AA9","AA9");

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\findcell.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod()]
        public void WorksheetSetStylesTest()
        {
            List<string> values = new List<string>(new string[] { "alpha", "beta", "charlie", "delta" });
            List<string> values2 = new List<string>(new string[] { "echo", "foxtrot", "golf", "hotel" });

            //Write to the spreadsheet
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");
            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);

            writer.PasteValues("A2", values, CellValues.String);
            writer.PasteValues("A3", values2, CellValues.String);

            //The centre four styles should be aligned to center
            SpreadsheetStyle style = SpreadsheetReader.GetDefaultStyle(doc);
            style.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            writer.SetStyle(style, "B2", "C3");

            //Set style in non existing cells
            writer.SetStyle(style, "B5", "C6");

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\stylerange.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod()]
        public void WorksheetMultipleStyleTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");

            SpreadsheetStyle style = SpreadsheetReader.GetDefaultStyle(doc);
            style.IsItalic = true;
            style.SetBackgroundColor("FF0000");//(red)

            WorksheetWriter.PasteText(doc, worksheetPart, "B", 2, "Hello world");
            WorksheetWriter.SetStyle(style, doc, worksheetPart, "B", 2);

            style = SpreadsheetReader.GetDefaultStyle(doc);
            style.IsBold = true;
            style.SetBackgroundColor("0000FF");//(blue)

            WorksheetWriter.SetStyle(style, doc, worksheetPart, "C", 3);
            WorksheetWriter.PasteText(doc, worksheetPart, "C", 3, "Hello world2");

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\styled2.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod()]
        public void WorksheetBorderTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");

            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);

            //Fill a background cell to make sure it is not overwritten
            SpreadsheetStyle style = SpreadsheetReader.GetDefaultStyle(doc);
            style.SetBackgroundColor("C0C0C0");
            //(grey)
            writer.SetStyle(style, "B2");

            writer.DrawBorder("B2", "D4", "FF0000", BorderStyleValues.Medium);

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\border.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod()]
        public void WorksheetOverlappingBorderTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");

            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);

            //Also use coordinates that are reversed
            writer.DrawBorder("B4", "D2", "FF0000", BorderStyleValues.Medium);
            writer.DrawBorder("E5", "C3", "0000FF", BorderStyleValues.Medium);

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\borderoverlap.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod()]
        public void WorksheetClearBorderTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");

            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);

            //Also use coordinates that are reversed
            writer.DrawBorder("B4", "D2", "FF0000", BorderStyleValues.Medium);
            writer.DrawBorder("E5", "C3", "0000FF", BorderStyleValues.Medium);

            //Now remove red
            writer.ClearBorder("B4", "D2");

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\borderclear.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod()]
        public void WorksheetInsertTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");

            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);

            writer.PasteNumber("D3", "2");
            writer.PasteNumber("D4", "3");
            writer.PasteNumber("D5", "4");

            //Add total without a calc chain
            writer.FindCell("D6").CellFormula = new CellFormula("SUM(D3:D5)");

            //Insert a row 
            writer.InsertRow(6);
            writer.PasteText("H6", "New content test");

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\insertrow.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod()]
        public void CompareFontTest()
        {
            Font x = new Font();
            Font y = new Font();
            FontSize size1 = new FontSize();
            FontSize size2 = new FontSize();
            Color color1 = new Color();
            Color color2 = new Color();
            FontName fontName1 = new FontName();
            FontName fontName2 = new FontName();
            FontFamily fontFamily1 = new FontFamily();
            FontFamily fontFamily2 = new FontFamily();
            FontScheme fontScheme1 = new FontScheme();
            FontScheme fontScheme2 = new FontScheme();

            color1.Rgb = "FFFF0000";
            color2.Rgb = "FFFF0000";
            size1.Val = 12;
            size2.Val = 12;
            fontName1.Val = "Calibri";
            fontName2.Val = "Calibri";
            fontFamily1.Val = 2;
            fontFamily2.Val = 2;
            fontScheme1.Val = FontSchemeValues.Minor;
            fontScheme2.Val = FontSchemeValues.Minor;

            x.AppendChild<Italic>(new Italic());
            x.AppendChild<FontSize>(size1);
            x.AppendChild<Color>(color1);
            x.AppendChild<FontName>(fontName1);
            x.AppendChild<FontFamily>(fontFamily1);
            x.AppendChild<FontScheme>(fontScheme1);

            y.AppendChild<Italic>(new Italic());
            y.AppendChild<FontSize>(size2);
            y.AppendChild<Color>(color2);
            y.AppendChild<FontName>(fontName2);
            y.AppendChild<FontFamily>(fontFamily2);
            y.AppendChild<FontScheme>(fontScheme2);

            //check they match
            Assert.IsTrue(SpreadsheetStyle.CompareFont(x, y), "Equal fonts do not compare.");

            //change a value
            size2.Val = 13;
            Assert.IsFalse(SpreadsheetStyle.CompareFont(x, y), "Unequal fonts do compare.");
        }

        [TestMethod()]
        public void WorksheetMergeCellsTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");

            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);
            SpreadsheetStyle style = SpreadsheetReader.GetDefaultStyle(doc);

            Cell cell = writer.PasteSharedText("B2", "Merged cells");
            style.IsUnderline = true;
            writer.SetStyle(style, "B2");

            Cell cell2 = writer.FindCell("C2");
            cell2.StyleIndex = cell.StyleIndex;

            writer.MergeCells("B2", "C2");

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\merge.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod()]
        public void WorksheetPrintAreaTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");
            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);

            writer.PasteText("A1", "Set print area to A1:B9");

            //Test setting the print area
            DefinedName area = writer.SetPrintArea("Sheet1", "A1", "B9");

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\printarea.xlsx", GetOutputFolder()), stream);

            //Assert.IsTrue(area != null, "Print area reference not returned.");
        }

        [TestMethod()]
        public void WorksheetDeleteRowTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");

            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);

            writer.PasteText("D3", "Row 3");
            writer.PasteText("D4", "Row 4");
            writer.PasteText("D5", "Row 5");

            //Delete row 3
            writer.DeleteRow(3);

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\deleterow.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod()]
        public void WorksheetDeleteRowsTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");

            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);

            writer.PasteText("D3", "Row 3");
            writer.PasteText("D4", "Row 4");
            writer.PasteText("D5", "Row 5");

            //Delete row 3
            writer.DeleteRows(3,2);

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\deleterows.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod()]
        public void WorksheetColumnWidthTest()
        {
            MemoryStream stream = SpreadsheetReader.Create();
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");

            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);

            writer.FindColumn("A").Width = 20;
            writer.FindColumn(2).Width = 20;
            writer.SetColumnWidth("C", 20);

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\columnwidths.xlsx", GetOutputFolder()), stream);
        }

        [TestMethod(), DeploymentItem("Templates\\columnstemplate.xlsx")]
        public void WorksheetColumnSplitTest()
        {
            MemoryStream stream = SpreadsheetReader.Copy(string.Format("{0}\\Templates\\columnstemplate.xlsx", Directory.GetCurrentDirectory()));
            SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(doc, "Sheet1");

            WorksheetWriter writer = new WorksheetWriter(doc, worksheetPart);

            writer.FindColumn("C").Width = 20;

            //Save to the memory stream, and then to a file
            SpreadsheetWriter.Save(doc);
            SpreadsheetWriter.StreamToFile(string.Format("{0}\\columnsplit.xlsx", GetOutputFolder()), stream);
        }

        private string GetOutputFolder()
        {
            return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        }

    }
}
