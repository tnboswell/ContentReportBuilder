using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using OfficeOpenXml;

namespace ContentReportBuilder.EPPlus.Builder.Excel.Sections
{
    public class WorksheetBuilder : DocumentSectionBuilderBase<ExcelWorksheet>, IDocumentSectionBuilder<ExcelWorksheet>
    {
        public WorksheetBuilder() : base(new DocumentElementFactory())
        {
        }
    }
}
