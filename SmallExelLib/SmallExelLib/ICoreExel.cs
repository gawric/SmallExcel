using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SmallExelLib.model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmallExelLib
{
    public interface ICoreExel
    {
        void setStyleDocument(Stylesheet style);
        SpreadsheetDocument GetDocument(string path);
        WorkbookPart  GetWorkBook(SpreadsheetDocument document);

        void setSheets(Sheets sheets);
        void CreateRootSheet(UInt32Value idSheet, DataModelSheet dmh, WorkbookPart workbookPart, string nameSheet);
        void AddNextSheet(UInt32Value idSheet, DataModelSheet dmh, WorkbookPart workbookPart, string nameSheet);

        void SaveWorkBookPart(WorkbookPart workbookPart);

        void CloseDocument(SpreadsheetDocument document);


    }
}
