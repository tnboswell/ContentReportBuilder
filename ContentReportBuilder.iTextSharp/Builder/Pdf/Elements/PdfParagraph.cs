using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using iTextSharp.text;

namespace ContentReportBuilder.ITextSharp.Builder.Pdf.Elements
{
    public class PdfParagraph : DocumentElementBuilder<Document>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Paragraph;

        protected override void Build(DocumentModelElement element, Document documentSection)
        {
            ParagraphData elementData = (ParagraphData)element.GetData();
            var paragraph = new Paragraph(elementData.Text);
            ParagraphFormating(element.WordStyle, paragraph);
            documentSection.Add(paragraph);
        }
    }
}
