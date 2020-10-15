using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.Styles;
using iTextSharp.text;
using System;
using System.Text;

namespace ContentReportBuilder.ITextSharp.Builder.Pdf.Elements
{
    public class PdfHeading : DocumentElementBuilder<Document>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Heading;

        protected override void Build(DocumentModelElement element, Document documentSection)
        {
            HeadingData elementData = (HeadingData)element.GetData();
            var paragraph = new Paragraph(elementData.Text);
            ParagraphFormating(element.WordStyle, paragraph);
            documentSection.Add(paragraph);
        }

       
    }
}
