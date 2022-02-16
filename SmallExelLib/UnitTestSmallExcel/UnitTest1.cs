using System;
using System.Collections.Generic;
using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SmallExelLib;
using SmallExelLib.model;
using SmallExelLib.variable;

namespace UnitTestSmallExcel
{
    [TestClass]
    public class UnitExcelCore
    {
        [TestMethod]
        public void TestStartCreateExcel()
        {
            ICoreExel coreExel = new CoreExel();
            SpreadsheetDocument document = coreExel.GetDocument("D:\\document.xlsx");
            WorkbookPart workbookPart = coreExel.GetWorkBook(document);
            Sheets sheets = null;

            coreExel.setStyleDocument(StatVariable.GenerateStyleSheet());
            coreExel.CreateRootSheet(1, GetDMS(), workbookPart, "Стены");
            coreExel.AddNextSheet(2, GetDMSOther(), workbookPart, "Перекрытия");
            

            coreExel.SaveWorkBookPart(workbookPart);
            coreExel.CloseDocument(document);

        }


        private DataModelSheet GetDMS()
        {
           return new DataModelSheet(GetColumnList() , GetHeaderColumnList() , GetRowsColumnList());
        }
        private DataModelSheet GetDMSOther()
        {
            return new DataModelSheet(GetColumnList(), GetHeaderColumnListOtherSheet(), GetRowsColumnList());
        }
        private List<ModelColumn> GetColumnList()
        {
            List<ModelColumn> list = new List<ModelColumn>();
            list.Add(new ModelColumn(1, 10, 10, true));
            list.Add(new ModelColumn(2, 10, 25, true));
            list.Add(new ModelColumn(3, 10, 40, true));
            list.Add(new ModelColumn(4, 10, 30, true));
            list.Add(new ModelColumn(5, 10, 20, true));
            list.Add(new ModelColumn(6, 10, 20, true));
            list.Add(new ModelColumn(7, 10, 20, true));
            list.Add(new ModelColumn(8, 10, 20, true));
            list.Add(new ModelColumn(9, 10, 20, true));
            list.Add(new ModelColumn(10, 10, 50, true));


            return list;
        }
        // Обычно всегда 1 т.к это первая ячейка
        //Last Element номер стиля, задавали в самом начале
        private List<ModelHeaderColumn> GetHeaderColumnList()
        {
            List<ModelHeaderColumn> list = new List<ModelHeaderColumn>();
            list.Add(new ModelHeaderColumn(1, 1, "No", CellValues.String, 1));
            list.Add(new ModelHeaderColumn(1, 2, "Классификатор", CellValues.String, 1));
            list.Add(new ModelHeaderColumn(1, 3, "Наим. Конструкций перекрытия", CellValues.String, 1));
            list.Add(new ModelHeaderColumn(1, 4, "Базовый Уровень", CellValues.String, 1));
            list.Add(new ModelHeaderColumn(1, 5, "Длина, м", CellValues.String, 1));
            list.Add(new ModelHeaderColumn(1, 6, "Ширина, м", CellValues.String, 1));
            list.Add(new ModelHeaderColumn(1, 7, "Площадь, м2", CellValues.String, 1));
            list.Add(new ModelHeaderColumn(1, 8, "Кол-во", CellValues.String, 1));
            list.Add(new ModelHeaderColumn(1, 9, "Объем м3", CellValues.String, 1));
            list.Add(new ModelHeaderColumn(1, 10, "Материал", CellValues.String, 1));

            return list;
        }

        // Обычно всегда 1 т.к это первая ячейка
        //Last Element номер стиля, задавали в самом начале
        private List<ModelHeaderColumn> GetHeaderColumnListOtherSheet()
        {
            List<ModelHeaderColumn> list = new List<ModelHeaderColumn>();
            list.Add(new ModelHeaderColumn(1, 1, "Новый 12", CellValues.String, 5));
            list.Add(new ModelHeaderColumn(1, 2, "Новый 23", CellValues.String, 5));
            list.Add(new ModelHeaderColumn(1, 3, "Новый 34", CellValues.String, 5));
            list.Add(new ModelHeaderColumn(1, 4, "Новый 4", CellValues.String, 5));
            list.Add(new ModelHeaderColumn(1, 5, "Новый 5", CellValues.String, 5));
            list.Add(new ModelHeaderColumn(1, 6, "Новый 6", CellValues.String, 5));
            list.Add(new ModelHeaderColumn(1, 7, "Новый 7", CellValues.String, 5));

            return list;
        }


        private List<List<ModelRows>> GetRowsColumnList()
        {
            List<List<ModelRows>> container = new List<List<ModelRows>>();
            int indexRows = 1;
            List<ModelRows> list1 = new List<ModelRows>();
            list1.Add(new ModelRows( 1, indexRows++.ToString(), CellValues.Number, 2));
            list1.Add(new ModelRows( 2, "8.2.4", CellValues.String, 2));
            list1.Add(new ModelRows( 3, "test1", CellValues.String, 2));
            list1.Add(new ModelRows( 4, "Этаж 12_+34,560 ", CellValues.String, 2));
            list1.Add(new ModelRows( 5, "0", CellValues.String, 2));
            list1.Add(new ModelRows( 6, "180", CellValues.String, 2));
            list1.Add(new ModelRows( 7, "366,7329037", CellValues.String, 2));
            list1.Add(new ModelRows( 8, "1", CellValues.String, 2));
            list1.Add(new ModelRows( 9, "216.5745494", CellValues.String, 2));
            list1.Add(new ModelRows( 10, "test4", CellValues.String, 2));


            List<ModelRows> list2 = new List<ModelRows>();
            list2.Add(new ModelRows( 1, indexRows++.ToString(), CellValues.Number, 2));
            list2.Add(new ModelRows( 2, "8.2.4", CellValues.String, 2));
            list2.Add(new ModelRows( 3, "test2", CellValues.String, 2));
            list2.Add(new ModelRows( 4, "Этаж 12_+34,500", CellValues.String, 2));
            list2.Add(new ModelRows( 5, "0", CellValues.String, 2));
            list2.Add(new ModelRows( 6, "200", CellValues.String, 2));
            list2.Add(new ModelRows( 7, "23,03476912", CellValues.String, 2));
            list2.Add(new ModelRows( 8, "1", CellValues.String, 2));
            list2.Add(new ModelRows( 9, "2721,018997", CellValues.String, 2));
            list2.Add(new ModelRows( 10, "test3", CellValues.String, 2));

            container.Add(list1);
            container.Add(list2);


            List<ModelRows> list3Itogo = new List<ModelRows>();
            list3Itogo.Add(new ModelRows(1, "", CellValues.Number, 1));
            list3Itogo.Add(new ModelRows(2, "Выше нуля", CellValues.String, 1));
            list3Itogo.Add(new ModelRows(3, "", CellValues.String, 1));
            list3Itogo.Add(new ModelRows(4, "Количество 76", CellValues.String, 1));
            list3Itogo.Add(new ModelRows(5, "", CellValues.String, 1));
            list3Itogo.Add(new ModelRows(6, "", CellValues.String, 1));
            list3Itogo.Add(new ModelRows(7, "", CellValues.String, 1));
            list3Itogo.Add(new ModelRows(8, "", CellValues.String, 1));
            list3Itogo.Add(new ModelRows(9, "", CellValues.String, 1));
            list3Itogo.Add(new ModelRows(10, "", CellValues.String, 1));

            List<ModelRows> list4Itogo = new List<ModelRows>();
            list4Itogo.Add(new ModelRows(1, "", CellValues.Number, 1));
            list4Itogo.Add(new ModelRows(2, "Ниже нуля", CellValues.String, 1));
            list4Itogo.Add(new ModelRows(3, "", CellValues.String, 1));
            list4Itogo.Add(new ModelRows(4, "Количество 0", CellValues.String, 1));
            list4Itogo.Add(new ModelRows(5, "", CellValues.String, 1));
            list4Itogo.Add(new ModelRows(6, "", CellValues.String, 1));
            list4Itogo.Add(new ModelRows(7, "", CellValues.String, 1));
            list4Itogo.Add(new ModelRows(8, "", CellValues.String, 1));
            list4Itogo.Add(new ModelRows(9, "", CellValues.String, 1));
            list4Itogo.Add(new ModelRows(10, "", CellValues.String, 1));

            List<ModelRows> list5Itogo = new List<ModelRows>();
            list5Itogo.Add(new ModelRows(1, "", CellValues.Number, 1));
            list5Itogo.Add(new ModelRows(2, "Итого", CellValues.String, 1));
            list5Itogo.Add(new ModelRows(3, "", CellValues.String, 1));
            list5Itogo.Add(new ModelRows(4, "", CellValues.String, 1));
            list5Itogo.Add(new ModelRows(5, "", CellValues.String, 1));
            list5Itogo.Add(new ModelRows(6, "", CellValues.String, 1));
            list5Itogo.Add(new ModelRows(7, "", CellValues.String, 1));
            list5Itogo.Add(new ModelRows(8, "", CellValues.String, 1));
            list5Itogo.Add(new ModelRows(9, "4502,846", CellValues.String, 1));
            list5Itogo.Add(new ModelRows(10, "", CellValues.String, 1));

            container.Add(list3Itogo);
            container.Add(list4Itogo);
            container.Add(list5Itogo);


            return container;
        }

     }
}
