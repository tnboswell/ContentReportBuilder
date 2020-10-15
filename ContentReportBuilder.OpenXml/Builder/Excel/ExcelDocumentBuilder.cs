using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.OpenXml.Builder.Excel;
using ContentReportBuilder.OpenXml.Builder.Excel.Sections;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.IO;


namespace ContentReportBuilder.OpenXml.Builder.Excel
{
    public class ExcelDocumentBuilder : DocumentBuilderBase<Sheet, SpreadsheetDocument, WorkbookPart>, IDocumentBuilder
    {
        public override DocumentBuilderType Type => DocumentBuilderType.Excel;

        public ExcelDocumentBuilder() : base(new WorksheetBuilder())
        {

        }

        public override Stream BuildStream(DocumentModel document)
        {
            Stream mem = new MemoryStream();
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(mem, SpreadsheetDocumentType.Workbook))
            {
                Build(document, package);
                return mem;
            }
        }

        public override void BuildFile(DocumentModel document)
        {
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(Filename(document.MetaData), SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = Build(document, package);
                workbookPart.Workbook.Save();
            }
        }

        public override WorkbookPart Build(DocumentModel document, SpreadsheetDocument package)
        {
            WorkbookPart workbookPart = package.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

            UInt32Value sheetNumber = 1;
            foreach (var section in document.DocumentSections ?? new List<DocumentModelSection>())
            {
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = sheetNumber, Name = section.MetaData.Title };
                _sectionBuilder.Build(Type, section, sheet);
                sheets.Append(sheet);
            }
            return workbookPart;
        }
    }
}
