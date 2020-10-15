using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using Xceed.Words.NET;

namespace ContentReportBuilder.docX.Builder.Word.Elements
{
    public class WordParagraph : DocumentElementBuilder<DocX>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Paragraph;

        protected override void Build(DocumentModelElement element, DocX documentSection)
        {
            ParagraphData elementData = (ParagraphData)element.GetData();
            var heading = documentSection.InsertParagraph(elementData.Text);
            ParagraphFormating(element.WordStyle, heading);
        }
    }
}
