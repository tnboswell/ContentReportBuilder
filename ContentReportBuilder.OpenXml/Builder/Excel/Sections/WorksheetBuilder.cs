using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ContentReportBuilder.OpenXml.Builder.Excel.Sections
{
    public class WorksheetBuilder : DocumentSectionBuilderBase<Sheet>, IDocumentSectionBuilder<Sheet>
    {
        public WorksheetBuilder() : base(new DocumentElementFactory())
        {
        }
    }
}
