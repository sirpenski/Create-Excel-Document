// ********************************************************************
// Copyright Paul F. Sirpenski
// MIT License
// ********************************************************************

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using System.Text;
using Microsoft.AspNetCore.Mvc;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using CreateExcelDocument.Models;




namespace CreateExcelDocument.Controllers
{
    public class ExcelSampleController : Controller
    {

        // These are class level variables set in createstylesheet
        private int CELLSTYLE_DEFAULT = 0;
        private int HEADER_CELLSTYLE_LEFT_JUSTIFIED = 0;
        private int HEADER_CELLSTYLE_RIGHT_JUSTIFIED = 0;
        private int DATA_CELLSTYLE_TEXT = 0;
        private int DATA_CELLSTYLE_DATE = 0;
        private int DATA_CELLSTYLE_INVOICE_NUMBER = 0;
        private int DATA_CELLSTYLE_WILL_PICKUP = 0;
        private int DATA_CELLSTYLE_QTY = 0;
        private int DATA_CELLSTYLE_CURRENCY = 0;
        private int FOOTER_CELLSTYLE_TOTAL_CURRENCY = 0;
        private int FOOTER_CELLSTYLE_TOTAL_LABEL_TEXT = 0;


        // *****************************************************************************
        // *****************************************************************************
        // Index Default Action
        // *****************************************************************************
        // *****************************************************************************

        public IActionResult Index()
        {
            return View();
        }





        #region CompleteWorksheet


        // ******************************************************************************
        // ******************************************************************************
        // This action Creates an Excel File In Memory and Returns it as an ActionResult
        // ******************************************************************************
        // ******************************************************************************
        public IActionResult CreateExcelFile()
        {
            IActionResult rslt = new BadRequestResult();

            string WorksheetName = "CompleteWorksheet";


            // create a new memory stream;
            MemoryStream ms = new MemoryStream();


            // ---------------------------- BEGIN CONSTRUCTING WORKBOOK -------------------


            // boilerplate
            // create a new document to the stream
            SpreadsheetDocument doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook, false);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookPart = doc.AddWorkbookPart();

            // add the workbook to the workbookpart (1 workbook for workbookpart)
            workbookPart.Workbook = new Workbook();


            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            // add a new worksheet to the worksheet part.  Note the initialization 
            // of SheetData to the worksheet contructor
            // worksheetPart.Worksheet = new Worksheet(new SheetData());
            worksheetPart.Worksheet = new Worksheet();


            // ---------------------------- BEGIN ADDING STYLESHEET TO DOCUMENT -------------------

            // add a new style part to the workbook
            WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();

            // create the stylesheet and assign it to the stylepart.stylesheet property
            stylePart.Stylesheet = CreateStyleSheet();

            // save stylesheet to style part
            stylePart.Stylesheet.Save();


            // --------------------------- END ADDING STYLESHEET TO DOCUMENT ----------------------



            // ----------------------------BEGIN ADDING COLUMNS --------------- -------------------

            // create the columns
            Columns worksheetColumns = CreateWorksheetColumns();

            // append the columns.  NOTE!!! Only works if you provide nothing to the 
            // new Worksheet declaration
            worksheetPart.Worksheet.AppendChild(worksheetColumns);


            // save the workbook part
            workbookPart.Workbook.Save();


            // ------------------------------- END ADDING COLUMNS --------------------------------



            // ---------------------------- BEGIN PRELIMINARY BUILD OF WORKSHEET -----------------
            worksheetPart.Worksheet.Append(new SheetData());

            // Add a Sheets collection to the Workbook.
            Sheets sheets = doc.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // Create a new worksheet associate it with the workbook.
            Sheet sheet = new Sheet() { Id = doc.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = WorksheetName };

            // add the worksheet to the Sheets collection
            sheets.Append(sheet);

            // save the workbook part
            workbookPart.Workbook.Save();


            // Get the sheetData object.  This is actually the object you will be adding the data 
            // rows to.
            SheetData sheetData = (SheetData)worksheetPart.Worksheet.GetFirstChild<SheetData>();

            // --------------------------- END PRELIMINARY BUILD OF WORKSHEET -------------------


            // ----------------------------- BEGIN BUILDING THE CELLS (ROW BY ROW)---------------

            // first, do the header row
            Row rw = new Row();

            // append the row
            sheetData.Append(rw);

            // now we need to set the row index of the row for 
            // cell reference, formula purposes
            rw.RowIndex = UInt32Value.FromUInt32((UInt32)sheetData.Elements().Count());

            // create the row
            CreateHeaderDataRow(rw);

            // create  a dataset
            List<DataItem> dataItems = CreateDataSet();

            // lets set a row pointer to the beginning row
            int BeginningDataRow = 0;


            // now, create each data row by appending it to sheetdata
            for (int i = 0; i < dataItems.Count; i++)
            {
                

                // new data row
                rw = new Row();

                // append the row
                sheetData.Append(rw);

                // now we need to set the row index of the row for 
                // cell reference and formula purposes
                rw.RowIndex = UInt32Value.FromUInt32((UInt32)sheetData.ChildElements.Count);

                // create the data row
                CreateDataRow(dataItems[i], rw);

                // track the beginning row because we will use it to buil the total line
                if (i == 0)
                {
                    BeginningDataRow = sheetData.ChildElements.Count;
                }

            }

            // now, we build a total row
            rw = new Row();

            // append to the sheet data
            sheetData.Append(rw);

            // now we need to set the row index of the row for 
            // cell reference and formula purposes
            rw.RowIndex = UInt32Value.FromUInt32((UInt32)sheetData.ChildElements.Count);

            // create the total row.  We know the begin row is 1 because the header is at 0.  We know the 
            // we could track the beginning row easily enough but for 
            CreateFooterRow(BeginningDataRow, dataItems.Count, rw);


            // ------------------------------- END BUILDING THE DATA CELLS -----------------------


            // ------------------------------- BEGIN CLOSING OUT WORKBOOK ------------------------

            // save the worksheet
            worksheetPart.Worksheet.Save();

            // save the workbook
            workbookPart.Workbook.Save();


            // close the document and flush to stream
            doc.Close();

            // -------------------------------END CLOSING WORKBOOK ------------------------------



            // rewind the memory stream
            ms.Seek(0, SeekOrigin.Begin);

            // return the file stream
            rslt = new FileStreamResult(ms, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");


            // back to the browser
            return rslt;

        }

        #endregion



        #region CreateWorksheetColumns

        // ******************************************************************************
        // ******************************************************************************
        // CreateWorksheetColumns configures the sizing of columns that 
        // comprise the worksheet
        // ******************************************************************************
        // ******************************************************************************
        private Columns CreateWorksheetColumns()
        {
            // define a new columns object
            Columns workSheetColumns = new Columns();

            // invoice number column
            Column col = new Column();
            col.Width = DoubleValue.FromDouble(16.0);
            col.Min = UInt32Value.FromUInt32((UInt32)1);
            col.Max = col.Min;
            col.CustomWidth = BooleanValue.FromBoolean(true);
            workSheetColumns.Append(col);

            // date column
            col = new Column();
            col.Width = DoubleValue.FromDouble(25.0);
            col.Min = UInt32Value.FromUInt32((UInt32)2);
            col.Max = col.Min;
            col.CustomWidth = BooleanValue.FromBoolean(true);
            workSheetColumns.Append(col);


            // first name column
            col = new Column();
            col.Width = DoubleValue.FromDouble(20);
            col.Min = UInt32Value.FromUInt32((UInt32)3);
            col.Max = col.Min;
            col.CustomWidth = BooleanValue.FromBoolean(true);
            workSheetColumns.Append(col);

            // last name column
            col = new Column();
            col.Width = DoubleValue.FromDouble(20.0);
            col.Min = UInt32Value.FromUInt32((UInt32)4);
            col.Max = col.Min;
            col.CustomWidth = BooleanValue.FromBoolean(true);
            workSheetColumns.Append(col);

            // will pickup column
            col = new Column();
            col.Width = DoubleValue.FromDouble(15.0);
            col.Min = UInt32Value.FromUInt32((UInt32)5);
            col.Max = col.Min;
            col.CustomWidth = BooleanValue.FromBoolean(true);
            workSheetColumns.Append(col);


            // qty column
            col = new Column();
            col.Width = DoubleValue.FromDouble(15.0);
            col.Min = UInt32Value.FromUInt32((UInt32)6);
            col.Max = col.Min;
            col.CustomWidth = BooleanValue.FromBoolean(true);
            workSheetColumns.Append(col);

            // unit price column
            col = new Column();
            col.Width = DoubleValue.FromDouble(15.0);
            col.Min = UInt32Value.FromUInt32((UInt32)7);
            col.Max = col.Min;
            col.CustomWidth = BooleanValue.FromBoolean(true);
            workSheetColumns.Append(col);

            // subtotal column
            col = new Column();
            col.Width = DoubleValue.FromDouble(15.0);
            col.Min = UInt32Value.FromUInt32((UInt32)8);
            col.Max = col.Min;
            col.CustomWidth = BooleanValue.FromBoolean(true);
            workSheetColumns.Append(col);

            return workSheetColumns;
        }

        #endregion


        #region CreateWorksheetStyleSheet

        // ******************************************************************************
        // ******************************************************************************
        // CreateWorksheetStylesheet creates the stylesheet for the worksheet 
        //
        // When creating a stylesheet, the idea is to: 
        // First, Define all the number formats, fonts, fills, and borders as separate entities.
        // Then, pick a number format, a font, a fill, and a border and create a cell format.  
        // Repeat this for as many cell formats as required in the worksheet. It can be 
        // any amount.
        // ******************************************************************************
        // ******************************************************************************
        private Stylesheet CreateStyleSheet()
        {
            // values to easily keep track of entities so when be 
            // finally build the cell formats, we do it with something friendly
            int NUMBERING_FORMAT_DATETIME = 200;
            int NUMBERING_FORMAT_INVOICE_NUMBER = 201;
            int NUMBERING_FORMAT_QTY = 202;
            int NUMBERING_FORMAT_CURRENCY = 203;

            // we are going to set these when we build the fills, that way, 
            // we don't have to come back up here everytime and modify the values;
            int FONTID_DEFAULT = 0;
            int FONTID_DEFAULT_BOLD = 0;
            int FONTID_RED_BOLD = 0;
            
            int FILLID_DEFAULT = 0;
            int FILLID_PATTERN_VALUE_GRAY_125 = 0;
            int FILLID_PATTERN_GOLD = 0;
            int FILLID_PATTERN_GREEN = 0;

            int BORDERID_DEFAULT = 0;
            int BORDERID_GRAY = 1;


            Stylesheet styleSheet = new Stylesheet();

            // define a new numbering format collection.  this collection will hold all the 
            // numbering formats used throughout the worksheet(s)

            styleSheet.NumberingFormats = new NumberingFormats();

            // THis is for the date, we give the number format an ID of 200.  Note 
            // ID is a UInt32Value
            NumberingFormat numberingFormat = new NumberingFormat();
            numberingFormat.FormatCode = StringValue.FromString("mm/dd/yyyy hh:mm:ss");
            numberingFormat.NumberFormatId = UInt32Value.FromUInt32((UInt32)NUMBERING_FORMAT_DATETIME);
            styleSheet.NumberingFormats.Append(numberingFormat);

            // this number format is for the invoice number
            numberingFormat = new NumberingFormat();
            numberingFormat.FormatCode = StringValue.FromString("00000000");
            numberingFormat.NumberFormatId = UInt32Value.FromUInt32((UInt32)NUMBERING_FORMAT_INVOICE_NUMBER);
            styleSheet.NumberingFormats.Append(numberingFormat);

            // this number format is for the quantity
            numberingFormat = new NumberingFormat();
            numberingFormat.FormatCode = StringValue.FromString("#");
            numberingFormat.NumberFormatId = UInt32Value.FromUInt32((UInt32)NUMBERING_FORMAT_QTY);
            styleSheet.NumberingFormats.Append(numberingFormat);

            // this format is for the unit price and the subtotal columns.
            numberingFormat = new NumberingFormat();
            numberingFormat.FormatCode = StringValue.FromString("$#.00");
            numberingFormat.NumberFormatId = UInt32Value.FromUInt32((UInt32)NUMBERING_FORMAT_CURRENCY);
            styleSheet.NumberingFormats.Append(numberingFormat);

            // update the collection count.  Don't know why object library can't do this but it doesn't
            styleSheet.NumberingFormats.Count = UInt32Value.FromUInt32((UInt32)styleSheet.NumberingFormats.ChildElements.Count);




            // Define a new FOnts collection.  This collection will contain all the fonts 
            // used in the worksheet(s).  REMEMBER THE INDEXES on these.  O Based
            styleSheet.Fonts = new Fonts();

            // index 0
            Font font = new Font();         // Default font
            styleSheet.Fonts.Append(font);
            FONTID_DEFAULT = styleSheet.Fonts.ChildElements.Count - 1;



            // default font bold
            font = new Font();
            font.Bold = new Bold();
            font.Bold.Val = BooleanValue.FromBoolean(true);
            styleSheet.Fonts.Append(font);
            FONTID_DEFAULT_BOLD = styleSheet.Fonts.ChildElements.Count - 1;




            // index 2. Bold Face Red.  We will use this for the headers
            font = new Font();
            font.Bold = new Bold();
            font.Bold.Val = BooleanValue.FromBoolean(true);
            font.Color = new Color();
            font.Color.Rgb = HexBinaryValue.FromString("FF0000");
            styleSheet.Fonts.Append(font);
            FONTID_RED_BOLD = styleSheet.Fonts.ChildElements.Count - 1;


            // update the font collection count
            styleSheet.Fonts.Count = UInt32Value.FromUInt32((UInt32)styleSheet.Fonts.ChildElements.Count);





            // define a fills collection.  Fills are used to create the background and foreground colors.  
            // they use another object called the pattern fill.  NOTE!!!! you have to always define the 
            // two preset fills.
            styleSheet.Fills = new Fills();
            
            // Fill Index 0
            Fill fill = new Fill();
            PatternFill PatternFillPreset = new PatternFill();
            PatternFillPreset.PatternType = PatternValues.None;
            fill.PatternFill = PatternFillPreset;
            styleSheet.Fills.Append(fill);
            FILLID_DEFAULT = styleSheet.Fills.ChildElements.Count - 1;

            // Fill Index 1.  Defaults By Micorosoft
            fill = new Fill();
            PatternFillPreset = new PatternFill();
            PatternFillPreset.PatternType = PatternValues.Gray125;
            fill.PatternFill = PatternFillPreset;
            styleSheet.Fills.Append(fill);
            FILLID_PATTERN_VALUE_GRAY_125 = styleSheet.Fills.ChildElements.Count - 1;


            // Fill Index 2 (Custom - Gold)
            fill = new Fill();
            PatternFill patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.Solid;
            patternFill.ForegroundColor = new ForegroundColor();
            patternFill.ForegroundColor.Rgb = HexBinaryValue.FromString("f9df02");
            fill.PatternFill = patternFill;
            styleSheet.Fills.Append(fill);
            FILLID_PATTERN_GOLD = styleSheet.Fills.ChildElements.Count - 1;


            // Fill Index 3 (Custom - Green)
            fill = new Fill();
            patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.Solid;
            patternFill.ForegroundColor = new ForegroundColor();
            patternFill.ForegroundColor.Rgb = HexBinaryValue.FromString("00ff00");
            fill.PatternFill = patternFill;
            styleSheet.Fills.Append(fill);
            FILLID_PATTERN_GREEN = styleSheet.Fills.ChildElements.Count - 1;


            // update the fills collection count
            styleSheet.Fills.Count = UInt32Value.FromUInt32((UInt32)styleSheet.Fills.ChildElements.Count);




            // Define the borders used in the worksheets.  
            styleSheet.Borders = new Borders();

            // default border
            Border border = new Border();
            styleSheet.Borders.Append(border);
            BORDERID_DEFAULT = styleSheet.Borders.ChildElements.Count - 1;



            string BorderColorString = "b4b4b4";
            border = new Border();
            border.LeftBorder = new LeftBorder();
            border.LeftBorder.Style = BorderStyleValues.Thin;
            border.LeftBorder.Color = new Color();
            border.LeftBorder.Color.Rgb = HexBinaryValue.FromString(BorderColorString);
            border.RightBorder = new RightBorder();
            border.RightBorder.Style = BorderStyleValues.Thin;
            border.RightBorder.Color = new Color();
            border.RightBorder.Color.Rgb = HexBinaryValue.FromString(BorderColorString);
            border.BottomBorder = new BottomBorder();
            border.BottomBorder.Style = BorderStyleValues.Thin;
            border.BottomBorder.Color = new Color();
            border.BottomBorder.Color.Rgb = HexBinaryValue.FromString(BorderColorString);
            border.TopBorder = new TopBorder();
            border.TopBorder.Style = BorderStyleValues.Thin;
            border.TopBorder.Color = new Color();
            border.TopBorder.Color.Rgb = HexBinaryValue.FromString(BorderColorString);
            styleSheet.Borders.Append(border);
            BORDERID_GRAY = styleSheet.Borders.ChildElements.Count - 1;


            // update the borders collection count
            styleSheet.Borders.Count = UInt32Value.FromUInt32((UInt32)styleSheet.Borders.ChildElements.Count);




            // create a new cell formats collection for the stylesheet
            styleSheet.CellFormats = new CellFormats();

            // index 0 - DEfault Cell Format
            CellFormat cellFormat = new CellFormat();
            styleSheet.CellFormats.Append(cellFormat);
            CELLSTYLE_DEFAULT = styleSheet.CellFormats.ChildElements.Count - 1;


            // index 1 (Header For Left Justified Cells)
            cellFormat = new CellFormat();
            cellFormat.FontId = UInt32Value.FromUInt32((UInt32)FONTID_RED_BOLD);
            cellFormat.ApplyFont = BooleanValue.FromBoolean(true);
            cellFormat.FillId = UInt32Value.FromUInt32((UInt32)FILLID_PATTERN_GOLD);
            cellFormat.ApplyFill = BooleanValue.FromBoolean(true);
            cellFormat.BorderId = UInt32Value.FromUInt32((UInt32)BORDERID_GRAY);
            cellFormat.ApplyBorder = BooleanValue.FromBoolean(true);
            cellFormat.Alignment = new Alignment();
            cellFormat.Alignment.Horizontal = HorizontalAlignmentValues.Left;
            cellFormat.Alignment.Vertical = VerticalAlignmentValues.Top;
            cellFormat.ApplyAlignment = BooleanValue.FromBoolean(true);
            styleSheet.CellFormats.Append(cellFormat);
            HEADER_CELLSTYLE_LEFT_JUSTIFIED = styleSheet.CellFormats.ChildElements.Count - 1;


            // index 2 (Header For Right Justified Cells)
            cellFormat = new CellFormat();
            cellFormat.FontId = UInt32Value.FromUInt32((UInt32)FONTID_RED_BOLD);
            cellFormat.ApplyFont = BooleanValue.FromBoolean(true);
            cellFormat.FillId = UInt32Value.FromUInt32((UInt32)FILLID_PATTERN_GOLD);
            cellFormat.ApplyFill = BooleanValue.FromBoolean(true);
            cellFormat.BorderId = UInt32Value.FromUInt32((UInt32)BORDERID_GRAY);
            cellFormat.ApplyBorder = BooleanValue.FromBoolean(true);
            cellFormat.Alignment = new Alignment();
            cellFormat.Alignment.Horizontal = HorizontalAlignmentValues.Right;
            cellFormat.Alignment.Vertical = VerticalAlignmentValues.Top;
            cellFormat.ApplyAlignment = BooleanValue.FromBoolean(true);
            styleSheet.CellFormats.Append(cellFormat);
            HEADER_CELLSTYLE_RIGHT_JUSTIFIED = styleSheet.CellFormats.ChildElements.Count - 1;



            // index 3  TEXT CELLS
            cellFormat = new CellFormat();
            cellFormat.Alignment = new Alignment();
            cellFormat.Alignment.Horizontal = HorizontalAlignmentValues.Left;
            cellFormat.Alignment.Vertical = VerticalAlignmentValues.Top;
            cellFormat.ApplyAlignment = BooleanValue.FromBoolean(true);
            styleSheet.CellFormats.Append(cellFormat);
            DATA_CELLSTYLE_TEXT = styleSheet.CellFormats.ChildElements.Count - 1;


            // (Date)
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = UInt32Value.FromUInt32((UInt32)NUMBERING_FORMAT_DATETIME);
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormat.Alignment = new Alignment();
            cellFormat.Alignment.Horizontal = HorizontalAlignmentValues.Left;
            cellFormat.Alignment.Vertical = VerticalAlignmentValues.Top;
            cellFormat.ApplyAlignment = BooleanValue.FromBoolean(true);
            styleSheet.CellFormats.Append(cellFormat);
            DATA_CELLSTYLE_DATE = styleSheet.CellFormats.ChildElements.Count - 1;


            // Invoice Number
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = UInt32Value.FromUInt32((UInt32)NUMBERING_FORMAT_INVOICE_NUMBER);
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormat.Alignment = new Alignment();
            cellFormat.Alignment.Horizontal = HorizontalAlignmentValues.Left;
            cellFormat.Alignment.Vertical = VerticalAlignmentValues.Top;
            cellFormat.ApplyAlignment = BooleanValue.FromBoolean(true);
            styleSheet.CellFormats.Append(cellFormat);
            DATA_CELLSTYLE_INVOICE_NUMBER = styleSheet.CellFormats.ChildElements.Count - 1;


            // index 6  WILL PICKUP.  BOOLEAN CELL
            cellFormat = new CellFormat();
            cellFormat.Alignment = new Alignment();
            cellFormat.Alignment.Horizontal = HorizontalAlignmentValues.Left;
            cellFormat.Alignment.Vertical = VerticalAlignmentValues.Top;
            cellFormat.ApplyAlignment = BooleanValue.FromBoolean(true);
            styleSheet.CellFormats.Append(cellFormat);
            DATA_CELLSTYLE_WILL_PICKUP = styleSheet.CellFormats.ChildElements.Count - 1;


            // Qty
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = UInt32Value.FromUInt32((UInt32)NUMBERING_FORMAT_QTY);
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormat.Alignment = new Alignment();
            cellFormat.Alignment.Horizontal = HorizontalAlignmentValues.Right;
            cellFormat.Alignment.Vertical = VerticalAlignmentValues.Top;
            cellFormat.ApplyAlignment = BooleanValue.FromBoolean(true);
            styleSheet.CellFormats.Append(cellFormat);
            DATA_CELLSTYLE_QTY = styleSheet.CellFormats.ChildElements.Count - 1;



            // Unit, Total - Currency
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = UInt32Value.FromUInt32((UInt32)NUMBERING_FORMAT_CURRENCY);
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormat.Alignment = new Alignment();
            cellFormat.Alignment.Horizontal = HorizontalAlignmentValues.Right;
            cellFormat.Alignment.Vertical = VerticalAlignmentValues.Top;
            cellFormat.ApplyAlignment = BooleanValue.FromBoolean(true);
            styleSheet.CellFormats.Append(cellFormat);
            DATA_CELLSTYLE_CURRENCY = styleSheet.CellFormats.ChildElements.Count - 1;



            // Total Label
            cellFormat = new CellFormat();
            cellFormat.Alignment = new Alignment();
            cellFormat.Alignment.Horizontal = HorizontalAlignmentValues.Right;
            cellFormat.Alignment.Vertical = VerticalAlignmentValues.Top;
            cellFormat.ApplyAlignment = BooleanValue.FromBoolean(true);
            styleSheet.CellFormats.Append(cellFormat);
            FOOTER_CELLSTYLE_TOTAL_LABEL_TEXT = styleSheet.CellFormats.ChildElements.Count - 1;



            // Total All Cells  - Currency
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = UInt32Value.FromUInt32((UInt32)NUMBERING_FORMAT_CURRENCY);
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormat.FontId = UInt32Value.FromUInt32((UInt32)FONTID_DEFAULT_BOLD);
            cellFormat.ApplyFont = BooleanValue.FromBoolean(true);
            cellFormat.FillId = UInt32Value.FromUInt32((UInt32)FILLID_PATTERN_GREEN);
            cellFormat.ApplyFill = BooleanValue.FromBoolean(true);
            cellFormat.BorderId = UInt32Value.FromUInt32((UInt32)BORDERID_GRAY);
            cellFormat.ApplyBorder = BooleanValue.FromBoolean(true);
            cellFormat.Alignment = new Alignment();
            cellFormat.Alignment.Horizontal = HorizontalAlignmentValues.Right;
            cellFormat.Alignment.Vertical = VerticalAlignmentValues.Top;
            cellFormat.ApplyAlignment = BooleanValue.FromBoolean(true);
            styleSheet.CellFormats.Append(cellFormat);
            FOOTER_CELLSTYLE_TOTAL_CURRENCY = styleSheet.CellFormats.ChildElements.Count - 1;


            // now update th cell formats count
            styleSheet.CellFormats.Count = UInt32Value.FromUInt32((UInt32)styleSheet.CellFormats.ChildElements.Count);



            return styleSheet;

        }

        #endregion



        #region CreateHeaderDataRow

        // ******************************************************************************
        // ******************************************************************************
        // This gets the header data row.  It is just like any other data row however, 
        // this only has column headers values
        // ******************************************************************************
        // ******************************************************************************
        private void CreateHeaderDataRow(Row rw)
        {


            Cell c = new Cell();
            c.CellValue = new CellValue("INVOICE#");
            c.DataType = CellValues.String;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)HEADER_CELLSTYLE_LEFT_JUSTIFIED);
            c.CellReference = "A" + rw.RowIndex.ToString();
            rw.Append(c);


            DateTime cNow = DateTime.Now;
            c = new Cell();
            c.CellValue = new CellValue("DATE");
            c.DataType = CellValues.String;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)HEADER_CELLSTYLE_LEFT_JUSTIFIED);
            c.CellReference = "B" + rw.RowIndex.ToString();
            rw.Append(c);

            c = new Cell();
            c.CellValue = new CellValue("FIRST");
            c.DataType = CellValues.String;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)HEADER_CELLSTYLE_LEFT_JUSTIFIED);
            c.CellReference = "C" + rw.RowIndex.ToString();
            rw.Append(c);

            // last
            c = new Cell();
            c.CellValue = new CellValue("LAST");
            c.DataType = CellValues.String;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)HEADER_CELLSTYLE_LEFT_JUSTIFIED);
            c.CellReference = "D" + rw.RowIndex.ToString();
            rw.Append(c);

            // will pickup
            c = new Cell();
            c.CellValue = new CellValue("WILL PICKUP");
            c.DataType = CellValues.String;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)HEADER_CELLSTYLE_LEFT_JUSTIFIED);
            c.CellReference = "E" + rw.RowIndex.ToString();
            rw.Append(c);

            // qty header
            c = new Cell();
            c.CellValue = new CellValue("QTY");
            c.DataType = CellValues.String;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)HEADER_CELLSTYLE_RIGHT_JUSTIFIED);
            c.CellReference = "F" + rw.RowIndex.ToString();
            rw.Append(c);

            c = new Cell();
            c.CellValue = new CellValue("UNITPRICE");
            c.DataType = CellValues.String;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)HEADER_CELLSTYLE_RIGHT_JUSTIFIED);
            c.CellReference = "G" + rw.RowIndex.ToString();
            rw.Append(c);

            c = new Cell();
            c.CellValue = new CellValue("SUBTOTAL");
            c.DataType = CellValues.String;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)HEADER_CELLSTYLE_RIGHT_JUSTIFIED);
            c.CellReference = "H" + rw.RowIndex.ToString();
            rw.Append(c);

            return;

        }

        #endregion


        #region CreateDataRow

        // ******************************************************************************
        // ******************************************************************************
        // Create Data Row simply adds the cells to the row for each data record
        // ******************************************************************************
        // ******************************************************************************
        private void CreateDataRow(DataItem itm, Row rw)
        {

            // invoice number (A)
            Cell c = new Cell();
            c.CellValue = new CellValue(itm.InvoiceNumber.ToString());
            c.DataType = CellValues.Number;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)DATA_CELLSTYLE_INVOICE_NUMBER);
            c.CellReference = "A" + rw.RowIndex.ToString();
            rw.Append(c);

            // invoice date.  (B)
            c = new Cell();
            c.CellValue = new CellValue(itm.InvoiceDate);
            c.DataType = CellValues.Date;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)DATA_CELLSTYLE_DATE);
            c.CellReference = "B" + rw.RowIndex.ToString();
            rw.Append(c);

            // first name (C)
            c = new Cell();
            c.CellValue = new CellValue(itm.First);
            c.DataType = CellValues.String;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)DATA_CELLSTYLE_TEXT);
            c.CellReference = "C" + rw.RowIndex.ToString();
            rw.Append(c);


            // last name (D)
            c = new Cell();
            c.CellValue = new CellValue(itm.Last);
            c.DataType = CellValues.String;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)DATA_CELLSTYLE_TEXT);
            c.CellReference = "D" + rw.RowIndex.ToString();
            rw.Append(c);

            // will pickup (E).  
            // Note, boolean values are "1" (true), "0" (false)"  
            // Therefore, if you have a csharp boolean value, you will need to transpose it to 
            // one or zero.  We will do that here using the ternary operator.
            int convertedBooleanValue = itm.WillPickUp ? 1 : 0;
            c = new Cell();
            c.CellValue = new CellValue(convertedBooleanValue.ToString());
            c.DataType = CellValues.Boolean;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)DATA_CELLSTYLE_WILL_PICKUP);
            c.CellReference = "E" + rw.RowIndex.ToString();
            rw.Append(c);


            // qty (F)
            c = new Cell();
            c.CellValue = new CellValue(itm.Qty.ToString());
            c.DataType = CellValues.Number;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)DATA_CELLSTYLE_QTY);
            c.CellReference = "F" + rw.RowIndex.ToString();
            rw.Append(c);
            Cell QtyCell = c;



            // unit price (G)
            c = new Cell();
            c.CellValue = new CellValue(itm.UnitPrice.ToString());
            c.DataType = CellValues.Number;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)DATA_CELLSTYLE_CURRENCY);
            c.CellReference = "G" + rw.RowIndex.ToString();
            rw.Append(c);
            Cell UnitPriceCell = c;



            // build the cell formula object.  The formula is "=F1 * G1" 
            // assuming row 1
            CellFormula cellFormula = new CellFormula();
            cellFormula.Text = "=" + QtyCell.CellReference + "*" + UnitPriceCell.CellReference;

            // now build the cell.  NOTE, we don't assign a cell value because 
            // that will come from the formula
            c = new Cell();
            c.CellFormula = cellFormula;
            c.DataType = CellValues.Number;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)DATA_CELLSTYLE_CURRENCY);
            c.CellReference = "H" + rw.RowIndex.ToString();
            rw.Append(c);

        }

        #endregion




        #region CreateFooterRow

        // ******************************************************************************
        // ******************************************************************************
        // Creates THe Footer Row (total)
        // ******************************************************************************
        // ******************************************************************************
        private void CreateFooterRow(int BeginRow, int DataItemCount, Row rw)
        {

            // invoice number (A)
            Cell c = new Cell();
            c.CellValue = new CellValue("");
            c.DataType = CellValues.String;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)CELLSTYLE_DEFAULT);
            c.CellReference = "A" + rw.RowIndex.ToString();
            rw.Append(c);

            // invoice date.  (B)
            c = new Cell();
            c.CellValue = new CellValue("");
            c.DataType = CellValues.String;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)CELLSTYLE_DEFAULT);
            c.CellReference = "B" + rw.RowIndex.ToString();
            rw.Append(c);

            // first name (C)
            c = new Cell();
            c.CellValue = new CellValue("");
            c.DataType = CellValues.String;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)CELLSTYLE_DEFAULT);
            c.CellReference = "C" + rw.RowIndex.ToString();
            rw.Append(c);


            // last name (D)
            c = new Cell();
            c.CellValue = new CellValue("");
            c.DataType = CellValues.String;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)CELLSTYLE_DEFAULT);
            c.CellReference = "D" + rw.RowIndex.ToString();
            rw.Append(c);

  
            c = new Cell();
            c.CellValue = new CellValue("");
            c.DataType = CellValues.String;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)CELLSTYLE_DEFAULT);
            c.CellReference = "E" + rw.RowIndex.ToString();
            rw.Append(c);


            // qty (F)
            c = new Cell();
            c.CellValue = new CellValue("");
            c.DataType = CellValues.String;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)CELLSTYLE_DEFAULT);
            c.CellReference = "F" + rw.RowIndex.ToString();
            rw.Append(c);
            Cell QtyCell = c;



            // unit price (G)
            c = new Cell();
            c.CellValue = new CellValue("Total:");
            c.DataType = CellValues.String;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)FOOTER_CELLSTYLE_TOTAL_LABEL_TEXT);
            c.CellReference = "G" + rw.RowIndex.ToString();
            rw.Append(c);
            Cell UnitPriceCell = c;



            // build the cell formula object.  The formula is SUM(BEGINCELLREF : ENDCELLREF)
            string BeginCellRef = "H" + BeginRow.ToString();
            string EndCellRef = "H" + (BeginRow + DataItemCount - 1).ToString();
            CellFormula cellFormula = new CellFormula();
            cellFormula.Text = "=SUM(" + BeginCellRef + ":" + EndCellRef + ")";

            // now build the cell.  NOTE, we don't assign a cell value because 
            // that will come from the formula
            c = new Cell();
            c.CellFormula = cellFormula;
            c.DataType = CellValues.Number;
            c.StyleIndex = UInt32Value.FromUInt32((UInt32)FOOTER_CELLSTYLE_TOTAL_CURRENCY);
            c.CellReference = "H" + rw.RowIndex.ToString();
            rw.Append(c);

        }

        #endregion



        #region CreateDataSet

        // ******************************************************************************
        // ******************************************************************************
        // CreateDataSet simply creates a list of data records which is used to build the 
        // contents of the excel spreadsheet.
        // ******************************************************************************
        // ******************************************************************************
        private List<DataItem> CreateDataSet()
        {
            List<string> LASTNAMES = new List<string>() {"Johnson", "Earnhardt", "Gordon", "Petty", "Preece", "Logano", "Keselowski",
                                                        "Trump", "Obama", "Bush", "Clinton", "Reagan", "Ford", "Nixon" };
            List<string> FIRSTNAMES = new List<string>() { "Jimmy", "Dale", "Jeff", "Richard", "Ryan", "Joey", "Brad",
                                                            "Donald", "Barak", "George", "Bill", "Ronald", "Gerald", "Tricky"};
            List<DataItem> rt = new List<DataItem>();

            int LASTNAMES_COUNT = LASTNAMES.Count;
            int FIRSTNAMES_COUNT = FIRSTNAMES.Count;
            int ivn = 0;
            DateTime dt = DateTime.Now.AddDays(-365);
            dt = new DateTime(dt.Year, dt.Month, dt.Day, dt.Hour, dt.Minute, dt.Second);
            Random RND = new Random();

            ivn = RND.Next(1, 5000);

            for (int i = 0; i < 100; i++)
            {

                DataItem itm = new DataItem();
                itm.Last = LASTNAMES[RND.Next(0, LASTNAMES_COUNT)];
                itm.First = FIRSTNAMES[RND.Next(0, FIRSTNAMES_COUNT)];
                itm.InvoiceNumber = ivn + i;

                itm.WillPickUp = (RND.Next(0, 100) > 49) ? true : false;

                itm.Qty = (decimal)RND.Next(1, 100);
                itm.UnitPrice = (decimal)RND.Next(11, 39);
                itm.SubTotal = itm.Qty * itm.UnitPrice;
                itm.InvoiceDate = dt;

                if (i % 5 == 0)
                {
                    dt = dt.Add(new TimeSpan(1, 1, 2, 30));
                }
                else
                {
                    dt = dt.Add(new TimeSpan(2, 3, 35));
                }


                rt.Add(itm);
            }

            return rt;
        }


        #endregion



        #region DataItem Class

        // ******************************************************************************
        // ******************************************************************************
        // Data Item Class.
        // ******************************************************************************
        // ******************************************************************************

        private class DataItem
        {
            public string First { get; set; } = "";
            public string Last { get; set; } = "";
            public int InvoiceNumber { get; set; } = 0;
            public DateTime InvoiceDate { get; set; } = DateTime.MinValue;
            public Boolean WillPickUp { get; set; } = false;
            public decimal Qty { get; set; } = 0M;
            public decimal UnitPrice { get; set; } = 0M;
            public decimal SubTotal { get; set; } = 0;
        }


        #endregion



    

    }
}
