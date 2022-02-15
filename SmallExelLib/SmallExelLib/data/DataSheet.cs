using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SmallExelLib.model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SmallExelLib.data
{
    public class DataSheet
    {
        //Размер колонок 
        public void AddColumn(WorksheetPart workSheetPart, List<ModelColumn> listColumn)
        {
            // Задаем колонки и их ширину
            Columns lstColumns = workSheetPart.Worksheet.GetFirstChild<Columns>();
            Boolean needToInsertColumns = false;
            if (lstColumns == null)
            {
                lstColumns = new Columns();
                needToInsertColumns = true;
            }

            foreach (ModelColumn mColumn in listColumn)
            {
                lstColumns.Append(new Column() { Min = mColumn.Min, Max = mColumn.Max, Width = mColumn.Width, CustomWidth = mColumn.CustomWidth });
            }

            if (needToInsertColumns)
                workSheetPart.Worksheet.InsertAt(lstColumns, 0);
        }

        public void AddRowsHeader(SheetData sheetData, List<ModelHeaderColumn> listMHC)
        {
            //Добавим заголовки в первую строку
            Row row = new Row() { RowIndex = 1 };
            sheetData.Append(row);

            foreach (ModelHeaderColumn headeritem in listMHC)
            {
                InsertCell(row, headeritem.cell_num, headeritem.val, headeritem.type, headeritem.styleIndex);
            }


        }

        //indexRows - это номер строки
        public void AddRows(SheetData sheetData, List<List<ModelRows>> listContainer)
        {

            UInt32Value indexRow = 2;
          

            foreach(List<ModelRows> listMHC in listContainer)
            {
                // Добавляем в строку все стили подряд.
                Row row = new Row() { RowIndex = indexRow++ };
                sheetData.Append(row);

                foreach (ModelRows item in listMHC)
                {
                    InsertCell(row, item.cell_num, item.val, item.type, item.styleIndex);
                }
            }
            

        }


        //Добавление Ячейки в строку (На вход подаем: строку, номер колонки, тип значения, стиль)
        public void InsertCell(Row row, int cell_num, string val, CellValues type, uint styleIndex)
        {
            Cell refCell = null;
            Cell newCell = new Cell() { CellReference = cell_num.ToString() + ":" + row.RowIndex.ToString(), StyleIndex = styleIndex };
            row.InsertBefore(newCell, refCell);

            // Устанавливает тип значения.
            newCell.CellValue = new CellValue(val);
            newCell.DataType = new EnumValue<CellValues>(type);

        }

        //Важный метод, при вставки текстовых значений надо использовать.
        //Метод убирает из строки запрещенные спец символы.
        //Если не использовать, то при наличии в строке таких символов, вылетит ошибка.
        static string ReplaceHexadecimalSymbols(string txt)
        {
            string r = "[\x00-\x08\x0B\x0C\x0E-\x1F\x26]";
            return Regex.Replace(txt, r, "", RegexOptions.Compiled);
        }

        //Метод генерирует стили для ячеек (за основу взят код, найденный где-то в интернете)

    }
}
