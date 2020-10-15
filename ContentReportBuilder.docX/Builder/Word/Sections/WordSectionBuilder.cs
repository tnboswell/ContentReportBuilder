
using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using Xceed.Words.NET;

namespace ContentReportBuilder.docX.Builder.Word.Sections
{
    public class WordSectionBuilder : DocumentSectionBuilderBase<DocX>, IDocumentSectionBuilder<DocX>
    {
        public WordSectionBuilder() : base(new DocumentElementFactory())
        {
        }
    }
}
