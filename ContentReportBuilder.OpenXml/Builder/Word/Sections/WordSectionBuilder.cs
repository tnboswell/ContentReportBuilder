using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ContentReportBuilder.OpenXml.Builder.Word.Sections
{

    public class WordSectionBuilder : DocumentSectionBuilderBase<Body>, IDocumentSectionBuilder<Body>
    {
        public WordSectionBuilder() : base(new DocumentElementFactory())
        {
        }
    }
}
