Apache POI, взаимодействие с Excel
_________________________________________________________
Apache POI представляет собой API, который позволяет использовать файлы MS Office в Java приложениях. Данная библиотека разрабатывается и распространяется Apache Software Foundation и носит открытый характер. Apache POI включает классы и методы для чтения и записи информации в документы MS Office.

**Компоненты Apache POI**

**Описание компонентов**
- HSSF	Horrible Spreadsheet Format	Компонент чтения и записи файлов MS-Excel, формат XLS
- XSSF	XML Spreadsheet Format	Компонент чтения и записи файлов MS-Excel, формат XLSX
- HPSF	Horrible Property Set Format	Компонент получения наборов свойств файлов MS-Office
- HWPF	Horrible Word Processor Format	Компонент чтения и записи файлов MS-Word, формат DOC
- XWPF	XML Word Processor Format	Компонент чтения и записи файлов MS-Word, формат DOCX
- HSLF	Horrible Slide Layout Format	Компонент чтения и записи файлов PowerPoint, формат PPT
- XSLF	XML Slide Layout Format	Компонент чтения и записи файлов PowerPoint, формат PPTX
- HDGF	Horrible DiaGram Format	Компонент работы с файлами MS-Visio, формат VSD
- XDGF	XML DiaGram Format	Компонент работы с файлами MS-Visio, формат VSDX

Рассматриваются следующие классы, используемые для работы с файлами Excel из приложений Java.
- рабочая книга - HSSFWorkbook, XSSFWorkbook
- лист книги - HSSFSheet, XSSFSheet
- строка - HSSFRow, XSSFRow
- ячейка - HSSFCell, XSSFCell
- стиль - стили ячеек HSSFCellStyle, XSSFCellStyle
- шрифт - шрифт ячеек HSSFFont, XSSFFont

**Классы и методы Apache POI для работы с файлами Excel**

**Рабочая книга HSSFWorkbook, XSSFWorkbook**

_HSSFWorkbook
- org.apache.poi.hssf.usermodel
- класс чтения и записи файлов Microsoft Excel в формате .xls, совместим с версиями MS-Office 97-2003;

_XSSFWorkbook
- org.apache.poi.xssf.usermodel
- класс чтения и записи файлов Microsoft Excel в формате .xlsx, совместим с MS-Office 2007 или более поздней версии.

**Конструкторы класса HSSFWorkbook**

- HSSFWorkbook ();
- HSSFWorkbook (InternalWorkbook book);
- HSSFWorkbook (POIFSFileSystem  fs);
- HSSFWorkbook (NPOIFSFileSystem fs);
- HSSFWorkbook (POIFSFileSystem  fs, 
              boolean preserveNodes);
- HSSFWorkbook (DirectoryNode directory, 
              POIFSFileSystem fs, 
              boolean preserveNodes);
- HSSFWorkbook (DirectoryNode directory,
              boolean preserveNodes);
- HSSFWorkbook (InputStream s);
- HSSFWorkbook (InputStream s, 
              boolean preserveNodes);
    **preservenodes** является необязательным параметром, который определяет необходимость сохранения узлов типа макросы.
    
**Конструкторы класса XSSFWorkbook**
    
- XSSFWorkbook ();
// workbookType  создать .xlsx или .xlsm
- XSSFWorkbook (XSSFWorkbookType workbookType);
- XSSFWorkbook (OPCPackage   pkg );
- XSSFWorkbook (InputStream  is  );
- XSSFWorkbook (File         file);
- XSSFWorkbook (String       path);
