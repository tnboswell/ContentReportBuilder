using ContentReportBuilder.Builder;
using ContentReportBuilder.docX.Builder.Word.Elements;
using ContentReportBuilder.Models;
using System.Collections.Generic;
using Xceed.Document.NET;

namespace ContentReportBuilder.docX.Builder
{
    class DocumentElementFactory : DocumentElementFactoryBase
    {
        protected override Dictionary<DocumentElementType, IDocumentElementBuilder> ExcelElements()
        {
            return new Dictionary<DocumentElementType, IDocumentElementBuilder>();
        }

        protected override Dictionary<DocumentElementType, IDocumentElementBuilder> WordElements()
        {
            var elements = new Dictionary<DocumentElementType, IDocumentElementBuilder>();
            elements.Add(DocumentElementType.BarChart, new WordBarChart());
            elements.Add(DocumentElementType.Heading, new WordHeading());
            elements.Add(DocumentElementType.Image, new WordImage());
            elements.Add(DocumentElementType.LineChart, new WordLineChart());
            elements.Add(DocumentElementType.List, new WordList());
            elements.Add(DocumentElementType.Paragraph, new WordParagraph());
            elements.Add(DocumentElementType.PieChart, new WordPieChart());
            elements.Add(DocumentElementType.Table, new WordTable());
            elements.Add(DocumentElementType.Header, new WordHeader());
            elements.Add(DocumentElementType.Footer, new WordFooter());




            return elements;
        }
    }
}
