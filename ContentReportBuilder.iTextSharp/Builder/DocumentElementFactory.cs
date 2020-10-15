using ContentReportBuilder.Builder;
using ContentReportBuilder.ITextSharp.Builder.Pdf.Elements;
using ContentReportBuilder.Models;
using System.Collections.Generic;

namespace ContentReportBuilder.ITextSharp.Builder
{
    public class DocumentElementFactory : DocumentElementFactoryBase
    {

        protected override Dictionary<DocumentElementType, IDocumentElementBuilder> PdfElements()
        {
            var elements = new Dictionary<DocumentElementType, IDocumentElementBuilder>();
            elements.Add(DocumentElementType.BarChart, new PdfBarChart());
            elements.Add(DocumentElementType.Heading, new PdfHeading());
            elements.Add(DocumentElementType.Image, new PdfImage());
            elements.Add(DocumentElementType.LineChart, new PdfLineChart());
            elements.Add(DocumentElementType.List, new PdfList());
            elements.Add(DocumentElementType.Paragraph, new PdfParagraph());
            elements.Add(DocumentElementType.PieChart, new PdfPieChart());
            elements.Add(DocumentElementType.Table, new PdfTable());
            elements.Add(DocumentElementType.Header, new PdfHeader());
            elements.Add(DocumentElementType.Footer, new PdfFooter());
            return elements;
        }
    }
}
