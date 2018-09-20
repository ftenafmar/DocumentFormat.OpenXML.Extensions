using System;
using System.Xml;
using System.Data;
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
    public class WorksheetWriter
    {

        private SpreadsheetDocument _spreadsheet;
        private WorksheetPart _worksheetPart;

        public WorksheetWriter(SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart)
        {
            if (spreadsheet == null) throw new ArgumentNullException("spreadsheet");
            if (worksheetPart == null) throw new ArgumentNullException("worksheetPart");

            _spreadsheet = spreadsheet;
            _worksheetPart = worksheetPart;
        }

        ///<summary>
        ///Returns the current spreadsheet
        ///</summary> 
        public SpreadsheetDocument Spreadsheet
        {
            get { return _spreadsheet; }
        }

        ///<summary>
        ///Returns the current worksheet
        ///</summary> 
        public WorksheetPart Worksheet
        {
            get { return _worksheetPart; }
        }

        ///<summary>
        ///Given a worksheet, column and row creates or returns a cell
        ///</summary> 
        public Cell FindCell(string reference)
        {
            return FindCell(SpreadsheetReader.ColumnFromReference(reference), SpreadsheetReader.RowFromReference(reference), Worksheet);
        }

        ///<summary>
        ///Writes a shared string to the row and column specified.
        ///</summary>
        public Cell PasteSharedText(string reference, string text)
        {
            return PasteValue(Spreadsheet, Worksheet, SpreadsheetReader.ColumnFromReference(reference), SpreadsheetReader.RowFromReference(reference), text, CellValues.SharedString, null);
        }

        ///<summary>
        ///Writes a shared string to the row and column specified.
        ///</summary>
        public Cell PasteSharedText(string reference, string text, SpreadsheetStyle style)
        {
            return PasteValue(Spreadsheet, Worksheet, SpreadsheetReader.ColumnFromReference(reference), SpreadsheetReader.RowFromReference(reference), text, CellValues.SharedString, style);
        }

        ///<summary>
        ///Writes a shared string to the range specified.
        ///</summary>
        public Cell PasteSharedTextRange(string range, string text)
        {
            range = SpreadsheetReader.GetDefinedName(Spreadsheet, range).InnerText;
            return PasteSharedText(SpreadsheetReader.ReferenceFromRange(range), text);
        }

        ///<summary>
        ///Writes a shared string to the range specified.
        ///</summary>
        public Cell PasteSharedTextRange(string range, string text, SpreadsheetStyle style)
        {
            range = SpreadsheetReader.GetDefinedName(Spreadsheet, range).InnerText;
            return PasteSharedText(SpreadsheetReader.ReferenceFromRange(range), text, style);
        }

        ///<summary>
        ///Writes an inline string to the row and column specified.
        ///</summary>
        public Cell PasteText(string reference, string text)
        {
            return PasteValue(Spreadsheet, Worksheet, SpreadsheetReader.ColumnFromReference(reference), SpreadsheetReader.RowFromReference(reference), text, CellValues.String, null);
        }

        ///<summary>
        ///Writes an inline string to the row and column specified.
        ///</summary>
        public Cell PasteText(string reference, string text, SpreadsheetStyle style)
        {
            return PasteValue(Spreadsheet, Worksheet, SpreadsheetReader.ColumnFromReference(reference), SpreadsheetReader.RowFromReference(reference), text, CellValues.String, style);
        }

        ///<summary>
        ///Writes an inline string to the range specified.
        ///</summary>
        public Cell PasteTextRange(string range, string text)
        {
            range = SpreadsheetReader.GetDefinedName(Spreadsheet, range).InnerText;
            return PasteText(SpreadsheetReader.ReferenceFromRange(range), text);
        }

        ///<summary>
        ///Writes an inline string to the range specified.
        ///</summary>
        public Cell PasteTextRange(string range, string text, SpreadsheetStyle style)
        {
            range = SpreadsheetReader.GetDefinedName(Spreadsheet, range).InnerText;
            return PasteText(SpreadsheetReader.ReferenceFromRange(range), text, style);
        }

        ///<summary>
        ///Writes a number to the row and column specified.
        ///</summary>
        public Cell PasteNumber(string reference, string number)
        {
            return PasteValue(Spreadsheet, Worksheet, SpreadsheetReader.ColumnFromReference(reference), SpreadsheetReader.RowFromReference(reference), number, CellValues.Number, null);
        }

        ///<summary>
        ///Writes a number to the row and column specified.
        ///</summary>
        public Cell PasteNumber(string reference, string number, SpreadsheetStyle style)
        {
            return PasteValue(Spreadsheet, Worksheet, SpreadsheetReader.ColumnFromReference(reference), SpreadsheetReader.RowFromReference(reference), number, CellValues.Number, style);
        }

        ///<summary>
        ///Writes a number to the range specified.
        ///</summary>
        public Cell PasteNumberRange(string range, string number)
        {
            range = SpreadsheetReader.GetDefinedName(Spreadsheet, range).InnerText;
            return PasteNumber(SpreadsheetReader.ReferenceFromRange(range), number);
        }

        ///<summary>
        ///Writes a number to the range specified.
        ///</summary>
        public Cell PasteNumberRange(string range, string number, SpreadsheetStyle style)
        {
            range = SpreadsheetReader.GetDefinedName(Spreadsheet, range).InnerText;
            return PasteNumber(SpreadsheetReader.ReferenceFromRange(range), number, style);
        }

        ///<summary>
        ///Writes adate to the row and column specified.
        ///</summary>
        public Cell PasteDate(string reference, DateTime date)
        {
            return PasteDate(Spreadsheet, Worksheet, SpreadsheetReader.ColumnFromReference(reference), SpreadsheetReader.RowFromReference(reference), date);
        }

        ///<summary>
        ///Writes a date to the range specified.
        ///</summary>
        public Cell PasteDateRange(string range, DateTime date)
        {
            range = SpreadsheetReader.GetDefinedName(Spreadsheet, range).InnerText;
            return PasteDate(SpreadsheetReader.ReferenceFromRange(range), date);
        }

        ///<summary>
        ///Writes any value to the row and column specified.
        ///</summary>
        public Cell PasteValue(string reference, string value, CellValues type)
        {
            return PasteValue(Spreadsheet, Worksheet, SpreadsheetReader.ColumnFromReference(reference), SpreadsheetReader.RowFromReference(reference), value, type, null);
        }

        ///<summary>
        ///Writes any value to the row and column specified.
        ///</summary>
        public Cell PasteValue(string reference, string value, CellValues type, SpreadsheetStyle style)
        {
            return PasteValue(Spreadsheet, Worksheet, SpreadsheetReader.ColumnFromReference(reference), SpreadsheetReader.RowFromReference(reference), value, type, style);
        }

        ///<summary>
        ///Writes any value to the range specified.
        ///</summary>
        public Cell PasteValueRange(string range, string value, CellValues type)
        {
            range = SpreadsheetReader.GetDefinedName(Spreadsheet, range).InnerText;
            return PasteValue(SpreadsheetReader.ReferenceFromRange(range), value, type);
        }

        ///<summary>
        ///Writes any value to the range specified.
        ///</summary>
        public Cell PasteValueRange(string range, string value, CellValues type, SpreadsheetStyle style)
        {
            range = SpreadsheetReader.GetDefinedName(Spreadsheet, range).InnerText;
            return PasteValue(SpreadsheetReader.ReferenceFromRange(range), value, type, style);
        }

        ///<summary>
        ///Writes any value to the row and column specified.
        ///</summary>
        public string PasteValues(string reference, List<string> values, CellValues type)
        {
            return PasteValues(Spreadsheet, Worksheet, SpreadsheetReader.ColumnFromReference(reference), SpreadsheetReader.RowFromReference(reference), values, type, null);
        }

        /// <summary>
        /// Pastes a datatable into a worksheet at the location specified and returns the rowindex of the last row.
        /// </summary>
        /// <remarks>
        /// Includes column headers and all columns.
        /// </remarks>
        public uint PasteDataTable(DataTable dt, string reference)
        {
            return PasteDataTable(dt, Spreadsheet, Worksheet, SpreadsheetReader.ColumnFromReference(reference), SpreadsheetReader.RowFromReference(reference), null, null);
        }

        /// <summary>
        /// Pastes a datatable into a worksheet at the location specified and returns the rowindex of the last row.
        /// </summary>
        /// <remarks>
        /// Includes column headers and all columns.
        /// </remarks>
        public uint PasteDataTable(DataTable dt, string reference, SpreadsheetStyle style)
        {
            return PasteDataTable(dt, Spreadsheet, Worksheet, SpreadsheetReader.ColumnFromReference(reference), SpreadsheetReader.RowFromReference(reference), null, style);
        }

        /// <summary>
        /// Pastes a datatable into a worksheet at the location specified and returns the rowindex of the last row.
        /// </summary>
        /// <remarks>
        /// Includes column headers and columns from the data table specified in the columnNames list.
        /// </remarks>
        public uint PasteDataTable(DataTable dt, string reference, List<string> columnNames)
        {
            return PasteDataTable(dt, Spreadsheet, Worksheet, SpreadsheetReader.ColumnFromReference(reference), SpreadsheetReader.RowFromReference(reference), columnNames, null);
        }

        /// <summary>
        /// Pastes a datatable into a worksheet at the location specified and returns the rowindex of the last row.
        /// </summary>
        /// <remarks>
        /// Includes column headers and columns from the data table specified in the columnNames list.
        /// </remarks>
        public uint PasteDataTable(DataTable dt, string reference, List<string> columnNames, SpreadsheetStyle style)
        {
            return PasteDataTable(dt, Spreadsheet, Worksheet, SpreadsheetReader.ColumnFromReference(reference), SpreadsheetReader.RowFromReference(reference), columnNames, style);
        }

        /// <summary>
        /// Inserts a datatable into a worksheet at the location specified and returns the rowindex of the last row.
        /// </summary>
        /// <remarks>
        /// Sufficient rows are inserted to make space for the data table. Includes column headers and all columns.
        /// </remarks>
        public uint InsertDataTable(DataTable dt, string reference)
        {
            return InsertDataTable(dt, Spreadsheet, Worksheet, SpreadsheetReader.ColumnFromReference(reference), SpreadsheetReader.RowFromReference(reference), null, null);
        }

        /// <summary>
        /// Inserts a datatable into a worksheet at the location specified and returns the rowindex of the last row.
        /// </summary>
        /// <remarks>
        /// Sufficient rows are inserted to make space for the data table. Includes column headers and all columns.
        /// </remarks>
        public uint InsertDataTable(DataTable dt, string reference, SpreadsheetStyle style)
        {
            return InsertDataTable(dt, Spreadsheet, Worksheet, SpreadsheetReader.ColumnFromReference(reference), SpreadsheetReader.RowFromReference(reference), null, style);
        }

        /// <summary>
        /// Pastes a datatable into a worksheet at the location specified and returns the rowindex of the last row.
        /// </summary>
        /// <remarks>
        /// Sufficient rows are inserted to make spac for the data table. Includes column headers and all columns.
        /// </remarks>
        public uint InsertDataTable(DataTable dt, string reference, List<string> columnNames)
        {
            return InsertDataTable(dt, Spreadsheet, Worksheet, SpreadsheetReader.ColumnFromReference(reference), SpreadsheetReader.RowFromReference(reference), columnNames, null);
        }

        /// <summary>
        /// Pastes a datatable into a worksheet at the location specified and returns the rowindex of the last row.
        /// </summary>
        /// <remarks>
        /// Sufficient rows are inserted to make spac for the data table. Includes column headers and all columns.
        /// </remarks>
        public uint InsertDataTable(DataTable dt, string reference, List<string> columnNames, SpreadsheetStyle style)
        {
            return InsertDataTable(dt, Spreadsheet, Worksheet, SpreadsheetReader.ColumnFromReference(reference), SpreadsheetReader.RowFromReference(reference), columnNames, style);
        }

        ///<summary>
        ///Sets the font style and colour for a cell.
        ///</summary>
        public Cell SetStyle(SpreadsheetStyle style, string reference)
        {
            return SetStyle(style, Spreadsheet, Worksheet, SpreadsheetReader.ColumnFromReference(reference), SpreadsheetReader.RowFromReference(reference));
        }

        ///<summary>
        ///Sets style information for a range of cells.
        ///</summary>
        public void SetStyle(SpreadsheetStyle style, string startReference, string endReference)
        {
            SetStyle(style, Spreadsheet, Worksheet, SpreadsheetReader.ColumnFromReference(startReference), SpreadsheetReader.RowFromReference(startReference), SpreadsheetReader.ColumnFromReference(endReference), SpreadsheetReader.RowFromReference(endReference));
        }

        ///<summary>
        ///Inserts a new row into worksheet and updates all existing cell references.
        ///</summary> 
        ///<remarks>
        ///Formula references are not updated by this method.
        ///</remarks>
        public void InsertRow(uint rowIndex)
        {
            InsertRows(rowIndex, 1, Worksheet);
        }

        ///<summary>
        ///Inserts one or more rows into worksheet and updates all existing cell references. Returns the last row.
        ///</summary> 
        ///<remarks>
        ///Formula references are not updated by this method.
        ///</remarks>
        public void InsertRows(uint rowIndex, uint count)
        {
            InsertRows(rowIndex, count, Worksheet);
        }

        ///<summary>
        ///Delete a row into worksheet and updates all existing cell references.
        ///</summary> 
        ///<remarks>
        ///Formula references are not updated by this method.
        ///</remarks>
        public void DeleteRow(uint rowIndex)
        {
            DeleteRows(rowIndex, 1, Worksheet);
        }

        ///<summary>
        ///Delete one or more rows into worksheet and updates all existing cell references.
        ///</summary> 
        ///<remarks>
        ///Formula references are not updated by this method.
        ///</remarks>
        public void DeleteRows(uint rowIndex, uint count)
        {
            DeleteRows(rowIndex, count, Worksheet);
        }

        ///<summary>
        ///Draws a border around the area defined by the two cell references.
        ///</summary>
        public void DrawBorder(string startReference, string endReference, string rgb, BorderStyleValues borderStyle)
        {
            DrawBorder(SpreadsheetReader.ColumnFromReference(startReference), SpreadsheetReader.RowFromReference(startReference), SpreadsheetReader.ColumnFromReference(endReference), SpreadsheetReader.RowFromReference(endReference), rgb, borderStyle, Spreadsheet, Worksheet);
        }

        ///<summary>
        ///Clears the border around the area defined by the two cell references.
        ///</summary>
        public void ClearBorder(string startReference, string endReference)
        {
            ClearBorder(SpreadsheetReader.ColumnFromReference(startReference), SpreadsheetReader.RowFromReference(startReference), SpreadsheetReader.ColumnFromReference(endReference), SpreadsheetReader.RowFromReference(endReference), Spreadsheet, Worksheet);
        }

        ///<summary>
        ///Merges the cell area defined by the two references into one cell.
        ///</summary> 
        public void MergeCells(string startReference, string endReference)
        {
            MergeCells(SpreadsheetReader.ColumnFromReference(startReference), SpreadsheetReader.RowFromReference(startReference), SpreadsheetReader.ColumnFromReference(endReference), SpreadsheetReader.RowFromReference(endReference), Spreadsheet, Worksheet);
        }
     
        ///<summary>
        /// Sets the defined name representing the print area for a worksheet
        /// </summary>
        public DefinedName SetPrintArea(string sheetName, string startReference, string endReference)
        {
            return SetPrintArea(Spreadsheet, sheetName, SpreadsheetReader.ColumnFromReference(startReference), SpreadsheetReader.RowFromReference(startReference), SpreadsheetReader.ColumnFromReference(endReference), SpreadsheetReader.RowFromReference(endReference));
        }

        ///<summary>
        ///Saves this worksheet and all related document parts.
        ///</summary> 
        public void Save()
        {
            Save(Spreadsheet, Worksheet);
        }

        /// <summary>
        /// Returns a row from the row index provided.
        /// </summary>
        public Row FindRow(uint index)
        {
            return FindRow(Worksheet.Worksheet.GetFirstChild<SheetData>(), index);
        }

        /// <summary>
        /// Returns a column from the column index provided.
        /// </summary>
        public Column FindColumn(uint index)
        {
            return FindColumn(Worksheet, index);
        }

        /// <summary>
        /// Returns a column from the column name provided.
        /// </summary>
        public Column FindColumn(string column)
        {
            return FindColumn(Worksheet, Convert.ToUInt32(SpreadsheetReader.GetColumnIndex(column)));
        }

        /// <summary>
        /// Sets the width of a column by index.
        /// </summary>
        public void SetColumnWidth(uint index, double width)
        {
            SetColumnWidth(Worksheet, index, width);
        }

        /// <summary>
        /// Sets the width of a column by name.
        /// </summary>
        public void SetColumnWidth(string column, double width)
        {
            SetColumnWidth(Worksheet, Convert.ToUInt32(SpreadsheetReader.GetColumnIndex(column)), width);
        }

        //Shared Functions

        ///<summary>
        ///Given a spreadsheet reference and text, writes a shared string to the row and column specified.
        ///</summary>
        public static Cell PasteSharedText(SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart, string column, uint row, string text)
        {
            return PasteValue(spreadsheet, worksheetPart, column, row, text, CellValues.SharedString, null);
        }

        ///<summary>
        ///Given a spreadsheet reference and text, writes an inline to the row and column specified.
        ///</summary>
        public static Cell PasteText(SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart, string column, uint row, string text)
        {
            return PasteValue(spreadsheet, worksheetPart, column, row, text, CellValues.String, null);
        }

        ///<summary>
        ///Given a spreadsheet reference and text, writes a number to the row and column specified.
        ///</summary>
        public static Cell PasteNumber(SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart, string column, uint row, string number)
        {
            return PasteValue(spreadsheet, worksheetPart, column, row, number, CellValues.Number, null);
        }

        ///<summary>
        ///Given a spreadsheet reference and date, writes the date to a row and column specified.
        ///</summary>
        public static Cell PasteDate(SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart, string column, uint row, DateTime date)
        {
            string output = GetNumericDate(date);
            Cell cell = PasteValue(spreadsheet, worksheetPart, column, row, output, CellValues.Number, null);
            cell.StyleIndex = GetReservedStyleIndex(Convert.ToUInt32(14), SpreadsheetReader.GetWorkbookStyles(spreadsheet));

            return cell;
        }

        ///<summary>
        ///Given a spreadsheet reference and text, writes any value to the row and column specified.
        ///</summary>
        public static Cell PasteValue(SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart, string column, uint row, string value, CellValues type, SpreadsheetStyle style)
        {
            //Get the cell, or insert a new one
            Cell cell = FindCell(column, row, worksheetPart);

            //If shared text then get the SharedStringTablePart. 
            //Create one in Excel by adding text, saving, then removing the text again.
            if (type == CellValues.SharedString)
            {
                SharedStringTablePart shareStringPart = spreadsheet.WorkbookPart.SharedStringTablePart;
                if (shareStringPart == null) throw new ApplicationException("Template does not contain a shared string table.");

                // Insert the text into the SharedStringTablePart.
                uint index = SpreadsheetWriter.GetSharedStringItem(value, shareStringPart);

                // Set the value of cell.
                cell.CellValue = new CellValue(index.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            }
            else
            {
                cell.CellValue = new CellValue(value);
                cell.DataType = new EnumValue<CellValues>(type);
            }

            //Set the style of the cell
            if (style != null) cell.StyleIndex = GetStyleIndex(style, SpreadsheetReader.GetWorkbookStyles(spreadsheet));

            return cell;
        }

        ///<summary>
        ///Given a worksheet, column and row, returns an existing cell or creates a new one if it doesnt exist
        ///</summary> 
        public static Cell FindCell(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = (columnName + rowIndex.ToString());
            int columnIndex = SpreadsheetReader.GetColumnIndex(columnName);
            
            //If the worksheet does not contain a row with the specified row index, insert one.
            Row row = FindRow(sheetData, rowIndex);

            //If there is not a cell with the specified column name, insert one.  
            IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference);
            if ((cells.Count() > 0))
            {
                return cells.First();
            }
            else
            {
                //Check the numerical value of the column portion of the cell reference.
                //Because the cells are in order, we add the new cell directly before first cell that is greater
                Cell refCell = null;
                
                foreach (Cell cell in row.Elements<Cell>())
                {
                    int colId = SpreadsheetReader.GetColumnIndex(SpreadsheetReader.ColumnFromReference(cell.CellReference.Value));
                    if (colId > columnIndex)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell();
                newCell.CellReference = cellReference;

                row.InsertBefore(newCell, refCell);

                return newCell;
            }
        }

        ///<summary>
        ///Inserts a new row into worksheet and updates all existing cell references.
        ///</summary> 
        ///<remarks>
        ///Formula references are not updated by this method.
        ///</remarks>
        public static void InsertRow(uint rowIndex, WorksheetPart worksheetPart)
        {
            InsertRows(rowIndex, 1, worksheetPart);
        }

        ///<summary>
        ///Inserts one or more rows into worksheet and updates all existing cell references. Returns the last row.
        ///</summary> 
        ///<remarks>
        ///Formula references are not updated by this method.
        ///</remarks>
        public static void InsertRows(uint rowIndex, uint count, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            //Get all the rows which are equal or greater than this row index
            IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex.Value >= rowIndex);

            //Move the cell references down by the number of rows
            foreach (Row row in rows)
            {
                row.RowIndex.Value += count;

                IEnumerable<Cell> cells = row.Elements<Cell>();
                foreach (Cell cell in cells)
                {
                    cell.CellReference = SpreadsheetReader.ColumnFromReference(cell.CellReference) + row.RowIndex.Value.ToString();
                }
            }
        }

        ///<summary>
        ///Delete a row into worksheet and updates all existing cell references.
        ///</summary> 
        ///<remarks>
        ///Formula references are not updated by this method.
        ///</remarks>
        public static void DeleteRow(uint rowIndex, WorksheetPart worksheetPart)
        {
            DeleteRows(rowIndex, 1, worksheetPart);
        }

        ///<summary>
        ///Delete one or more rows into worksheet and updates all existing cell references. Returns the last row.
        ///</summary> 
        ///<remarks>
        ///Formula references are not updated by this method.
        ///</remarks>
        public static void DeleteRows(uint rowIndex, uint count, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            //Remove rows to delete
            foreach (var row in sheetData.Elements<Row>().
                Where(r => r.RowIndex.Value >= rowIndex && r.RowIndex.Value < rowIndex + count).ToList())
            {
                row.Remove();
            }

            //Get all the rows which are equal or greater than this row index + row count
            IEnumerable<Row> rows = sheetData.Elements<Row>().Where(r => r.RowIndex.Value >= rowIndex + count);

            //Move the cell references up by the number of rows
            foreach (Row row in rows)
            {
                row.RowIndex.Value -= count;

                IEnumerable<Cell> cells = row.Elements<Cell>();
                foreach (Cell cell in cells)
                {
                    cell.CellReference = SpreadsheetReader.ColumnFromReference(cell.CellReference)
                        + row.RowIndex.Value.ToString();
                }
            }
        }

        /// <summary>
        /// Inserts a datatable into a worksheet at the location specified and returns the rowindex of the last row.
        /// </summary>
        /// <remarks>
        /// Sufficient rows are inserted into the worksheet to contain the data. Includes column headers and all columns, or columns if columnNames is supplied.
        /// </remarks>
        public static uint InsertDataTable(DataTable dt, SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart, string column, uint rowIndex, List<string> columnNames, SpreadsheetStyle style)
        {
            InsertRows(rowIndex, Convert.ToUInt32(dt.Rows.Count), worksheetPart);
            return PasteDataTable(dt, spreadsheet, worksheetPart, column, rowIndex, columnNames, style);
        }

        /// <summary>
        /// Pastes a datatable into a worksheet at the location specified and returns the rowindex of the last row.
        /// </summary>
        /// <remarks>
        /// Includes column headers and all columns, or columns if columnNames is supplied.
        /// </remarks>
        public static uint PasteDataTable(DataTable dt, SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart, string column, uint rowIndex, List<string> columnNames, SpreadsheetStyle style)
        {
            WorkbookStylesPart styles = SpreadsheetReader.GetWorkbookStyles(spreadsheet);
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            List<Type> numericTypes = new List<Type>(new Type[] { typeof(short), typeof(int), typeof(long), typeof(float), typeof(double), typeof(decimal), typeof(ushort), typeof(uint), typeof(ulong) });

            foreach (System.Data.DataRow dataRow in dt.Rows)
            {
                string colString = column;

                //Setup first reference
                Cell refCell = null; 
                if (SpreadsheetReader.GetColumnIndex(colString) > 0) refCell = WorksheetReader.GetCell(SpreadsheetReader.GetColumnName(colString,-1), rowIndex, worksheetPart);
                
                Row excelRow = FindRow(sheetData, rowIndex);

                foreach (System.Data.DataColumn dataColumn in dt.Columns)
                {
                    //Filter out columns not needed, if supplied
                    if (columnNames == null || columnNames.Contains(dataColumn.ColumnName))
                    {
                        string cellReference = (colString + rowIndex.ToString());
                        Cell excelCell;

                        // If there is not a cell with the specified column name, insert one.  
                        IEnumerable<Cell> excelCells = excelRow.Elements<Cell>().Where(c => c.CellReference.Value == cellReference);
                        if (excelCells.Count() > 0)
                        {
                            excelCell = excelCells.First();
                            refCell = excelCell;
                        }
                        else
                        {
                            excelCell = new Cell();
                            excelCell.CellReference = cellReference;

                            refCell = excelRow.InsertAfter(excelCell, refCell);
                        }

                        //Set the value
                        if (dataColumn.DataType == typeof(DateTime))
                        {
                            DateTime value;
                            if (DateTime.TryParse(dataRow.ItemArray[dataColumn.Ordinal].ToString(), out value))
                            {
                                excelCell.CellValue = new CellValue(GetNumericDate(value));
                                excelCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                               
                                //Use a reserved number format , or modify the current style
                                if (style == null)
                                {
                                    excelCell.StyleIndex = GetReservedStyleIndex(Convert.ToUInt32(14), styles);
                                }
                                else
                                {
                                    SpreadsheetStyle clone = style.Clone() as SpreadsheetStyle;
                                    clone.FormatCode = "mm-dd-yy";
                                    excelCell.StyleIndex = GetStyleIndex(clone, styles);
                                }
                            }
                        }
                        else
                        {
                            object cellValue = dataRow.ItemArray[dataColumn.Ordinal];

                            if (numericTypes.Contains(dataColumn.DataType))
                            {
                                excelCell.CellValue = new CellValue(SpreadsheetWriter.ToXmlNumeric(cellValue));
                            }
                            else if (dataColumn.DataType == typeof(bool))
                            {
                                excelCell.CellValue = new CellValue(SpreadsheetWriter.ToXmlBoolean(cellValue));
                            }
                            else
                            {
                                excelCell.CellValue = new CellValue(cellValue.ToString());
                            }

                            if (numericTypes.Contains(dataColumn.DataType))
                            {
                                excelCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            }
                            else if (dataColumn.DataType == typeof(bool))
                            {
                                excelCell.DataType = new EnumValue<CellValues>(CellValues.Boolean); //requires string type
                            }
                            else
                            {
                                excelCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            }

                            //Set the style
                            if (style != null) excelCell.StyleIndex = GetStyleIndex(style, styles);
                        }

                        //Get the next column
                        colString = SpreadsheetReader.GetColumnName(colString, 1);
                    }
                }

                rowIndex += Convert.ToUInt32(1);
            }

            return rowIndex;
        }

        ///<summary>
        ///Writes a list of values to the location specified, returning the ending column.
        ///</summary>
        public static string PasteValues(SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart, string column, uint rowIndex, List<string> values, CellValues type, SpreadsheetStyle style)
        {
            WorkbookStylesPart styles = SpreadsheetReader.GetWorkbookStyles(spreadsheet);
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            string colString = column;
            Cell refCell = WorksheetReader.GetCell(colString, rowIndex, worksheetPart);
            Row excelRow = FindRow(sheetData, rowIndex);

            //Check if the row is created
            // If the worksheet does not contain a row with the specified row index, insert one.
            if ((sheetData.Elements<Row>().Where(r => r.RowIndex.Value == rowIndex).Count() != 0))
            {
                excelRow = sheetData.Elements<Row>().Where(r => r.RowIndex.Value == rowIndex).First();
            }
            else
            {
                excelRow = new Row();
                excelRow.RowIndex = rowIndex;
                sheetData.Append(excelRow);
            }

            foreach (string value in values)
            {
                string cellReference = (colString + rowIndex.ToString());
                Cell excelCell;

                //If there is not a cell with the specified column name, insert one.  
                IEnumerable<Cell> excelCells = excelRow.Elements<Cell>().Where(c => c.CellReference.Value == cellReference);
                if ((excelCells.Count() > 0))
                {
                    excelCell = excelCells.First();
                }
                else
                {

                    excelCell = new Cell();
                    excelCell.CellReference = cellReference;

                    refCell = excelRow.InsertAfter(excelCell, refCell);
                }

                //Set the value
                excelCell.CellValue = new CellValue(value);
                excelCell.DataType = new EnumValue<CellValues>(type);

                //Merge the style with the current style
                if (style != null) excelCell.StyleIndex = GetStyleIndex(style, styles);

                //Get the next column
                colString = SpreadsheetReader.GetColumnName(colString, 1);
            }

            return colString;
        }

        ///<summary>
        ///Sets the style information for a cell.
        ///</summary>
        public static Cell SetStyle(SpreadsheetStyle style, SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart, string column, uint rowIndex)
        {
            WorkbookStylesPart styles = SpreadsheetReader.GetWorkbookStyles(spreadsheet);
            Cell cell = WorksheetWriter.FindCell(column, rowIndex, worksheetPart);
            //Get the cell, create if necessary

            cell.StyleIndex = GetStyleIndex(style, styles);

            return cell;
        }

        ///<summary>
        ///Sets the style for the defined area of cells.
        ///</summary>
        public static void SetStyle(SpreadsheetStyle style, SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart, string startColumn, uint startRowIndex, string endColumn, uint endRowIndex)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            WorkbookStylesPart styles = SpreadsheetReader.GetWorkbookStyles(spreadsheet);

            uint rowStart = startRowIndex;
            uint rowEnd = endRowIndex;
            string colStart = startColumn;
            string colEnd = endColumn;

            //Normalise the start and end columns so that the start is always smaller than the end
            if (rowEnd < rowStart)
            {
                rowStart = endRowIndex;
                rowEnd = startRowIndex;
            }

            if (string.Compare(colEnd, colStart) < 0)
            {
                colStart = endColumn;
                colEnd = startColumn;
            }

            //Loop through each row
            for (uint rowIndex = rowStart; rowIndex <= rowEnd; rowIndex++)
            {
                Row row = FindRow(sheetData, rowIndex);

                string col = string.Empty;

                //Loop through each column
                while (col != colEnd)
                {

                    if (col == string.Empty)
                    {
                        col = colStart;
                    }
                    else
                    {
                        col = SpreadsheetReader.GetColumnName(col, 1);
                    }

                    Cell cell = null;
                    string cellReference = string.Format("{0}{1}", col, rowIndex.ToString());

                    //Get the cell, or create if it doesnt exist
                    IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference == cellReference);
                    if ((cells.Count() > 0))
                    {
                        cell = cells.First();
                    }
                    else
                    {
                        // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                        Cell refCell = null;
                        foreach (Cell cellLoop in row.Elements<Cell>())
                        {
                            if (string.Compare(cellLoop.CellReference.Value, cellReference, true) > 0)
                            {
                                refCell = cell;
                                break;
                            }
                        }

                        Cell newCell = new Cell();
                        newCell.CellReference = cellReference;

                        row.InsertBefore(newCell, refCell);

                        cell = newCell;
                    }

                    cell.StyleIndex = GetStyleIndex(style, styles);
                }
            }
        }

        ///<summary>
        ///Returns an existing style index or creates a new style index from the style information provided.
        ///</summary>
        public static uint GetStyleIndex(SpreadsheetStyle style, WorkbookStylesPart styles)
        {
            //Find a CellFormat with the selected indexes, or create a new one
            UInt32 fontIndex = SpreadsheetWriter.CreateFont(style, styles);
            UInt32 fillIndex = SpreadsheetWriter.CreateFill(style, styles);
            UInt32 borderIndex = SpreadsheetWriter.CreateBorder(style, styles);

            NumberingFormat numberFormat = style.ToNumberFormat();
            UInt32 numberFormatIndex = 0;

            Alignment alignment = style.ToAlignment();
            
            //Lookup number format if required
            if (numberFormat != null) numberFormatIndex = SpreadsheetWriter.CreateNumberFormat(style, styles);

            uint index = 0;

            //Check all the existing cell formats
            foreach (CellFormat cellFormat in styles.Stylesheet.CellFormats)
            {
                //Compare indexed properties - Null check added by Mark Stevens
                if (cellFormat.FontId != null && cellFormat.FillId != null && cellFormat.BorderId != null && cellFormat.NumberFormatId != null)
                {
                    if (cellFormat.FontId.Value == fontIndex && cellFormat.FillId.Value == fillIndex && cellFormat.BorderId.Value == borderIndex && cellFormat.NumberFormatId.Value == numberFormatIndex)
                    {
                        if (SpreadsheetStyle.CompareAlignment(alignment, cellFormat.Alignment)) return index; //Compare direct properties
                    }
                }
                index += Convert.ToUInt32(1);
            }

            //Create a new cell format
            CellFormat newFormat = new CellFormat();

            newFormat.NumberFormatId = new UInt32Value(numberFormatIndex);
            newFormat.FontId = new UInt32Value(fontIndex);
            newFormat.FillId = new UInt32Value(fillIndex);
            newFormat.BorderId = new UInt32Value(borderIndex);
            newFormat.Alignment = alignment;

            newFormat.ApplyFont = new BooleanValue(true);
            if (newFormat.NumberFormatId.Value > 0) newFormat.ApplyNumberFormat = true;
            if (newFormat.FillId.Value > 0) newFormat.ApplyFill = true;
            if (newFormat.BorderId.Value > 0) newFormat.ApplyBorder = true;
            if (newFormat.Alignment != null) newFormat.ApplyAlignment = true;

            styles.Stylesheet.CellFormats.AppendChild<CellFormat>(newFormat);
            styles.Stylesheet.CellFormats.Count = index + Convert.ToUInt32(1);
            return index;

        }

        ///<summary>
        ///Returns an existing reserved number format style index or creates a new style index from the number format id provided.
        ///</summary>
        public static uint GetReservedStyleIndex(UInt32 numberFormatIndex, WorkbookStylesPart styles)
        {
            uint index = 0;

            //Check all the existing cell formats
            foreach (CellFormat cellFormat in styles.Stylesheet.CellFormats)
            {
                //Compare indexed properties
                if (cellFormat.NumberFormatId.Value == numberFormatIndex) return index;
                index += Convert.ToUInt32(1);
            }

            //Create a new cell format
            CellFormat newFormat = new CellFormat();

            newFormat.NumberFormatId = new UInt32Value(numberFormatIndex);
            newFormat.ApplyNumberFormat = true;

            styles.Stylesheet.CellFormats.AppendChild<CellFormat>(newFormat);
            styles.Stylesheet.CellFormats.Count = index + Convert.ToUInt32(1);
            return index;
        }

        ///<summary>
        ///Draws a border around the area defined by the two cell references.
        ///</summary>
        public static void DrawBorder(string startColumn, uint startRowIndex, string endColumn, uint endRowIndex, string rgb, BorderStyleValues borderStyle, SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart)
        {

            SpreadsheetStyle style = null;
            uint rowStart = startRowIndex;
            uint rowEnd = endRowIndex;
            string colStart = startColumn;
            string colEnd = endColumn;

            //Normalise the start and end columns so that the start is always smaller than the end
            if (rowEnd < rowStart)
            {
                rowStart = endRowIndex;
                rowEnd = startRowIndex;
            }

            if (string.Compare(colEnd, colStart, true) < 0)
            {
                colStart = endColumn;
                colEnd = startColumn;
            }

            //Loop through each row
            for (uint rowIndex = rowStart; rowIndex <= rowEnd; rowIndex++)
            {

                //Draw a line along the top or bottom
                if (rowIndex == rowStart || rowIndex == rowEnd)
                {
                    string col = string.Empty;

                    while (col != colEnd)
                    {
                        if (col == string.Empty)
                        {
                            col = colStart;
                        }
                        else
                        {
                            col = SpreadsheetReader.GetColumnName(col, 1);
                        }

                        //Get the style as the current cell
                        style = WorksheetReader.GetStyleWithDefault(spreadsheet, worksheetPart, col, rowIndex);

                        if (rowIndex == rowStart) style.SetBorderTop(rgb, borderStyle);
                        if (rowIndex == rowEnd) style.SetBorderBottom(rgb, borderStyle);

                        SetStyle(style, spreadsheet, worksheetPart, col, rowIndex);
                    }

                }

                //Add left border
                style = WorksheetReader.GetStyleWithDefault(spreadsheet, worksheetPart, colStart, rowIndex);
                style.SetBorderLeft(rgb, borderStyle);
                SetStyle(style, spreadsheet, worksheetPart, colStart, rowIndex);

                //Add right border
                style = WorksheetReader.GetStyleWithDefault(spreadsheet, worksheetPart, colEnd, rowIndex);
                style.SetBorderRight(rgb, borderStyle);
                SetStyle(style, spreadsheet, worksheetPart, colEnd, rowIndex);
            }
        }

        ///<summary>
        ///Draws a border around the area defined by the two cell references.
        ///</summary>
        public static void ClearBorder(string startColumn, uint startRowIndex, string endColumn, uint endRowIndex, SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart)
        {
            SpreadsheetStyle style = null;
            uint rowStart = startRowIndex;
            uint rowEnd = endRowIndex;
            string colStart = startColumn;
            string colEnd = endColumn;

            //Normalise the start and end columns so that the start is always smaller than the end
            if (rowEnd < rowStart)
            {
                rowStart = endRowIndex;
                rowEnd = startRowIndex;
            }

            if (string.Compare(colEnd, colStart, true) < 0)
            {
                colStart = endColumn;
                colEnd = startColumn;
            }

            //Loop through each row
            for (uint rowIndex = rowStart; rowIndex <= rowEnd; rowIndex++)
            {
                //Draw a line along the top or bottom
                if (rowIndex == rowStart || rowIndex == rowEnd)
                {
                    string col = string.Empty;

                    while (col != colEnd)
                    {
                        if (col == string.Empty)
                        {
                            col = colStart;
                        }
                        else
                        {
                            col = SpreadsheetReader.GetColumnName(col, 1);
                        }

                        //Get the style as the current cell
                        style = WorksheetReader.GetStyleWithDefault(spreadsheet, worksheetPart, col, rowIndex);

                        if (rowIndex == rowStart) style.ClearBorderTop();
                        if (rowIndex == rowEnd) style.ClearBorderBottom();

                        SetStyle(style, spreadsheet, worksheetPart, col, rowIndex);
                    }
                }

                //Add left border
                style = WorksheetReader.GetStyleWithDefault(spreadsheet, worksheetPart, colStart, rowIndex);
                style.ClearBorderLeft();
                SetStyle(style, spreadsheet, worksheetPart, colStart, rowIndex);

                //Add right border
                style = WorksheetReader.GetStyleWithDefault(spreadsheet, worksheetPart, colEnd, rowIndex);
                style.ClearBorderRight();
                SetStyle(style, spreadsheet, worksheetPart, colEnd, rowIndex);
            }
        }

        ///<summary>
        ///Creates a merged cell from the references supplied
        ///</summary>
        public static MergeCell MergeCells(string startColumn, uint startRowIndex, string endColumn, uint endRowIndex, SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart)
        {
            string reference = string.Format("{0}{1}:{2}{3}", new object[] { startColumn, startRowIndex, endColumn, endRowIndex }).ToUpper();
            Worksheet worksheet = worksheetPart.Worksheet;
            MergeCells cells = worksheet.GetFirstChild<MergeCells>();
            uint index = 0;

            //Create merge cells if doesnt exists, else look for existing 
            if (cells == null)
            {
                cells = new MergeCells();

                //The MergeCells element has to come directly after sheet data
                worksheet.InsertAfter(cells, worksheet.Elements<SheetData>().First());
            }
            else
            {
                foreach (MergeCell mergeCell in cells)
                {
                    if (mergeCell.Reference.Value.ToUpper() == reference) return mergeCell;
                    index += Convert.ToUInt32(1);
                }
            }

            //Add new merge cell
            var newMergeCell = new MergeCell();
            newMergeCell.Reference = reference;

            cells.AppendChild<MergeCell>(newMergeCell);
            cells.Count = index + Convert.ToUInt32(1);

            return new MergeCell();
        }

        /// <summary>
        /// Saves the worksheet and all related document parts.
        /// </summary>
        public static void Save(SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart)
        {
            SpreadsheetWriter.SetRowSpans(worksheetPart);
            SpreadsheetWriter.SetWorksheetDimension(worksheetPart);
            worksheetPart.Worksheet.Save();

            //Save the style information
            WorkbookStylesPart styles = SpreadsheetReader.GetWorkbookStyles(spreadsheet);
            styles.Stylesheet.Save();

            //Save the shared string table part
            if ((spreadsheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0))
            {
                SharedStringTablePart shareStringPart = spreadsheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                shareStringPart.SharedStringTable.Save();
            }

            //Save the workbook
            spreadsheet.WorkbookPart.Workbook.Save();
        }

        /// <summary>
        /// Gets the row specified at the row index, or creates a new row if one does not exist.
        /// </summary>
        public static Row FindRow(SheetData sheetData, uint rowIndex)
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
                row = new Row();
                row.RowIndex = index;

                //Get the position in the array to insert the row and insert it there
		        int count = 0;
		        foreach (var rowLoop in sheetData.Elements<Row>())
                {
			        if (rowLoop.RowIndex.Value > rowIndex) 
                    {
				        sheetData.InsertAt(row, count);
				        return row;
			        }
			        count += 1;
		        }

                sheetData.Append(row);
            }

            return row;
        }

        /// <summary>
        /// Gets a column from the sheet data
        /// </summary>
        public static Column FindColumn(WorksheetPart worksheetPart, uint columnIndex)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            Columns cols = worksheet.GetFirstChild<Columns>();
            Column col = null;

            //Add columns if dont exists
            if (cols == null)
            {
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                cols = new Columns();
                worksheet.InsertBefore<Columns>(cols, sheetData);
            }

            //Make sure the row exists
            var match = cols.Elements<Column>().Where(c => columnIndex >= c.Min && columnIndex <= c.Max);

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
            }
            else
            {
                col = new Column();
                col.Min = columnIndex;
                col.Max = columnIndex;
                col.Width = 9.140625;

                //Find the column to insert after
                var beforeCols = cols.Elements<Column>().Where(c => c.Max < columnIndex);
                cols.InsertAt<Column>(col, beforeCols.Count<Column>());
            }

            return col;
        }

        ///<summary>
        /// Set the column width of a column
        /// </summary>
        public static void SetColumnWidth(WorksheetPart worksheetPart, uint columnIndex, double width)
        {
            var col = FindColumn(worksheetPart, columnIndex);
            col.Width = width; 
        }

        ///<summary>
        /// Sets the defined name representing the print area for a worksheet
        /// </summary>
        public static DefinedName SetPrintArea(SpreadsheetDocument spreadsheet, string sheetName, string startColumn, uint startRowIndex, string endColumn, uint endRowIndex)
        {
            DefinedNames definedNames = null;
            IEnumerable<DefinedNames> elements = spreadsheet.WorkbookPart.Workbook.Descendants<DefinedNames>();
            DefinedName printAreaName = null;
            UInt32Value sheetId = SpreadsheetReader.GetSheetId(spreadsheet, sheetName);
            WorksheetPart worksheetPart = SpreadsheetReader.GetWorksheetPartByName(spreadsheet, sheetName);

            //Create or retrieve the defined names section
            if (elements.Count() == 0)
            {
                definedNames = new DefinedNames();

                //Need to insert directly after sheets
                Sheets sheets = spreadsheet.WorkbookPart.Workbook.Descendants<Sheets>().First();
                spreadsheet.WorkbookPart.Workbook.InsertAfter(definedNames, sheets);
            }
            else
            {
                definedNames = elements.First();

                //Find an existing print area defined name
                foreach (DefinedName definedName in definedNames)
                {
                    if (definedName.Name == "_xlnm.Print_Area" && definedName.LocalSheetId.HasValue && definedName.LocalSheetId.Value == sheetId.Value - 1)
                    {
                        printAreaName = definedName;
                        break;
                    }
                }
            }

            string range = string.Format("{0}!${1}${2}:${3}${4}", sheetName, startColumn, startRowIndex, endColumn, endRowIndex);

            //Set existing
            if (printAreaName != null)
            {
                printAreaName.Text = range;
            }

            //Create new
            else
            {
                printAreaName = new DefinedName(range);
                printAreaName.LocalSheetId = sheetId.Value - Convert.ToUInt32(1);
                printAreaName.Name = "_xlnm.Print_Area";

                definedNames.AppendChild(printAreaName);
            }

            //Set the page settings if not in the sheet
            if (WorksheetReader.GetPageSetup(spreadsheet, worksheetPart) == null) SetPageSetup(spreadsheet, worksheetPart, 9, OrientationValues.Default);

            return printAreaName;
        }

        ///<summary>
        /// sets the page setup options for a worksheet
        /// </summary>
        public static void SetPageSetup(SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart, uint paperSize, OrientationValues orientation)
        {
            Worksheet workSheet = worksheetPart.Worksheet;
            PageSetup pageSetup = WorksheetReader.GetPageSetup(spreadsheet, worksheetPart);

            //Create the page setup element if applicable
            if (pageSetup == null)
            {
                pageSetup = new PageSetup();

                PageMargins pageMargins = workSheet.GetFirstChild<PageMargins>();
                workSheet.InsertAfter(pageSetup, pageMargins);
            }

            pageSetup.PaperSize = paperSize;
            pageSetup.Orientation = orientation;
        }

        public static string GetNumericDate(DateTime date)
        {
            TimeSpan result = date - new DateTime(1900, 1, 1);
            int days = result.Days + 2; //Time difference + 2 

            double totalSeconds = 24.0F * 3600.0F;
            double fraction = ((date.Hour * 3600) + (date.Minute * 60) + date.Second) / totalSeconds; // Convert to a fraction of seconds in a day

            return (Convert.ToSingle(days) + fraction).ToString();
        }
    }
}
