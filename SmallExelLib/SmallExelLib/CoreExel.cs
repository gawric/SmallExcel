using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SmallExelLib.data;
using SmallExelLib.model;
using SmallExelLib.variable;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmallExelLib
{
    public class CoreExel : ICoreExel
    {
        private Sheets sheets;
        private Stylesheet styledocument;

        public void CreateRootSheet(UInt32Value idSheet , DataModelSheet dmh , WorkbookPart workbookPart, string name)
        {
           
                DataSheet ds = new DataSheet();
                if(workbookPart.Workbook == null) workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                FileVersion fv = new FileVersion();
                fv.ApplicationName = "Microsoft Office Excel";
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                AddStyleToSheets(workbookPart);


                //CreateColumn
                ds.AddColumn(worksheetPart, dmh.columnList);

                sheets = CreateSheets(workbookPart);
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = idSheet, Name = name };
                sheets.Append(sheet);

                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                ds.AddRowsHeader(sheetData, dmh.headerColumnList);
                ds.AddRows(sheetData, dmh.rowsColumnList);

            
        }

        public void AddNextSheet(UInt32Value idSheet, DataModelSheet dmh, WorkbookPart workbookPart, string nameSheet)
        {
            //----------------------------------------------------------
            DataSheet ds = new DataSheet();
            WorksheetPart worksheetPart2 = workbookPart.AddNewPart<WorksheetPart>();
            SheetData sheetData2 = new SheetData();
            Worksheet workSheet2 = worksheetPart2.Worksheet = new Worksheet(sheetData2);

            AddStyleToSheets(workbookPart);

            if (sheets == null)
            {
                if(workbookPart.Workbook == null)
                {
                    workbookPart.Workbook = new Workbook();
                }
                sheets = CreateSheets(workbookPart);
            }

            //CreateColumn
            ds.AddColumn(worksheetPart2, dmh.columnList);



            Sheet sheet2 = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart2), SheetId = idSheet, Name = nameSheet };
            sheets.Append(sheet2);

            ds.AddRowsHeader(sheetData2, dmh.headerColumnList);
            ds.AddRows(sheetData2, dmh.rowsColumnList);

        }

        private Sheets CreateSheets(WorkbookPart workbookPart)
        {
            if(sheets == null) return workbookPart.Workbook.AppendChild(new Sheets());
            return sheets;
        }
        public SpreadsheetDocument GetDocument(string path)
        {
            return SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        }

        public WorkbookPart GetWorkBook(SpreadsheetDocument document)
        {
            return document.AddWorkbookPart();
        }

        public void setSheets(Sheets sheets)
        {
            this.sheets = sheets;
        }
        private void AddStyleToSheets(WorkbookPart workbookPart)
        {
            if(workbookPart.WorkbookStylesPart == null)
            {
                if (styledocument != null)
                {
                    WorkbookStylesPart wbsp = workbookPart.AddNewPart<WorkbookStylesPart>();
                    wbsp.Stylesheet = styledocument;
                    wbsp.Stylesheet.Save();
                }
            }
            
        }
        public void setStyleDocument(Stylesheet style)
        {
            styledocument = style;
        }

        public void SaveWorkBookPart(WorkbookPart workbookPart)
        {
            workbookPart.Workbook.Save();
        }

        public void CloseDocument(SpreadsheetDocument document)
        {
            document.Close();
        }
    }
}
