using System;
using System.Xml;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXML.Extensions;

namespace DocumentFormat.OpenXml.Extensions
{
    ///<summary>
    ///Provides the base functionality around Spreadsheets
    ///</summary>
    public class WorksheetReader
    {
        //Blocks the constructor as this is intended to be a static library
        private WorksheetReader()
        {

        }

        /// <summary>
        /// Gets the row specified at the row index if it exists
        /// </summary>
        public static Row GetRow(SheetData sheetData, uint rowIndex)
        {
            Row row = null;
            uint index = rowIndex;

            //Make sure the row exists
            var match = sheetData.Elements<Row>().Where(r => r.RowIndex.Value == index);

            if (match.Count() != 0)
            {
                row = match.First();
            }
            else
            {
                return null;
            }

            return row;
        }

        /// <summary>
        /// Gets a column from the sheet data
        /// </summary>
        public static Column GetColumn(WorksheetPart worksheetPart, uint columnIndex)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            Columns cols = worksheet.GetFirstChild<Columns>();
            Column col = null;

            //If no columns have been created, return null
            if (cols == null) return null;

            var match = cols.Elements<Column>().Where(c => columnIndex >= c.Min && columnIndex <= c.Max);

            //Split up range of columns if required
            if (match.Count() != 0)
            {
                col = match.First();

                //Insert new column range before
                if (col.Min < columnIndex)
                {
                    Column before = col.CloneElement<Column>();
                    before.Max = columnIndex - 1;
                    cols.InsertBefore<Column>(before, col);

                    col.Min = columnIndex;
                }

                //Insert new column range after
                if (col.Max > columnIndex)
                {
                    Column after = col.CloneElement<Column>();
                    after.Min = columnIndex + 1;
                    cols.InsertAfter<Column>(after, col);

                    col.Max = columnIndex;
                }

                return col;
            }
            else
            {
                return null;
            }
        }

        ///<summary>
        ///Returns a cell is one exists at the location supplied.
        ///</summary>
        /// <remarks>
        ///Use FindCell instead if a missing cell should be created.
        ///</remarks>
        public static Cell GetCell(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = (columnName + rowIndex.ToString());

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex.Value == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex.Value == rowIndex).First();
            }
            else
            {
                return null;
            }

            //If there is not a cell with the specified column name, return null 
            IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference);
            if (cells.Count() > 0)
            {
                return cells.First();
            }
            else
            {
                return null;
            }
        }

        ///<summary>
        ///Gets the font, background and border style settings for a cell.
        ///</summary>
        ///<remarks>
        ///If the cell does not exists, then a null reference (nothing) is returned
        ///</remarks>
        public static SpreadsheetStyle GetStyle(SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart, string column, uint rowIndex)
        {
            WorkbookStylesPart styles = SpreadsheetReader.GetWorkbookStyles(spreadsheet);
            Cell cell = WorksheetReader.GetCell(column, rowIndex, worksheetPart); //Get the cell if it exists

            if (cell == null || cell.StyleIndex == null) return null;

            CellFormat cellFormat = (CellFormat) styles.Stylesheet.CellFormats.ChildElements[(int) cell.StyleIndex.Value];
            Font font = (Font) styles.Stylesheet.Fonts.ChildElements[(int) cellFormat.FontId.Value];
            Fill fill = (Fill) styles.Stylesheet.Fills.ChildElements[(int) cellFormat.FillId.Value];
            Border border = (Border) styles.Stylesheet.Borders.ChildElements[(int) cellFormat.BorderId.Value];
            Alignment alignment = cellFormat.Alignment;

            NumberingFormat numberFormat = null;

            //Lookup the number format
            if (styles.Stylesheet.NumberingFormats != null)
            {
                foreach (var numberFormatElement in styles.Stylesheet.NumberingFormats)
                {
                    var formatLoop = (NumberingFormat) numberFormatElement;
                    if (formatLoop.NumberFormatId.HasValue && cellFormat.NumberFormatId.HasValue && formatLoop.NumberFormatId.Value == cellFormat.NumberFormatId.Value)
                    {
                        numberFormat = formatLoop;
                        break;
                    }
                }
            }

            return new SpreadsheetStyle(font, fill, border, alignment, numberFormat);
        }

        ///<summary>
        ///Gets the font, background and border style settings for a cell.
        ///</summary>
        ///<remarks>
        ///If the cell does not exists, then the default style is returned instead.
        ///</remarks>
        public static SpreadsheetStyle GetStyleWithDefault(SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart, string column, uint rowIndex)
        {
            SpreadsheetStyle result = GetStyle(spreadsheet, worksheetPart, column, rowIndex);
            if (result == null) result = SpreadsheetReader.GetDefaultStyle(spreadsheet);

            return result;
        }

        ///<summary>
        ///Gets the page setup values for a worksheet.
        ///</summary>
        public static PageSetup GetPageSetup(SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart)
        {
            return worksheetPart.Worksheet.GetFirstChild<PageSetup>();
        }


    }
}