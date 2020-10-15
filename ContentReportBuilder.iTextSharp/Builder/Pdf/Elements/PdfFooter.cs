using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using iTextSharp.text;

namespace ContentReportBuilder.ITextSharp.Builder.Pdf.Elements
{
    public class PdfFooter : DocumentElementBuilder<Document>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Footer;

        protected override void Build(DocumentModelElement element, Document documentSection)
        {
        }
    }
}
