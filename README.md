# SmallExcel
Маленькая библиотека создана с помощью openxml и Net Framework 4.7.2

Использование:

# создание класса core
 ICoreExel coreExel = new CoreExel();
 
# создание основного документа
 SpreadsheetDocument document = coreExel.GetDocument("путь к файлу");
 WorkbookPart workbookPart = coreExel.GetWorkBook(document);
 
# Добавление стилей для документа
 coreExel.setStyleDocument(StatVariable.GenerateStyleSheet());
 
# создание первой страницы
 coreExel.CreateRootSheet(1, GetDMS(), workbookPart, "Отчет1");
 
# создание второй страницы
 coreExel.AddNextSheet(2, GetDMSOther(), workbookPart, "Отчет2");

# Закрытие документа
 coreExel.SaveWorkBookPart(workbookPart);
 coreExel.CloseDocument(document);
 
 
# Доп.описание
  GetDMS() -> возвращает набор данных для столбцов - строк.
  использует модель DataModelSheet(List<ModelColumn> columnList, List<ModelHeaderColumn> headerColumnList, List<List<ModelRows>> rowsColumnList)
  
  columnList -> размеры колонок и их количество
  headerColumnList -> первая строчка где обычно записываются имена колонок (пример: Наименование)
  rowsColumnList -> основной массив данных для заполнения листа
  
# Примеры использования есть в классе: UnitExcelCore->TestStartCreateExcel