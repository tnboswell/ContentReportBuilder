using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.OpenXml.Builder.Excel.Elements;
using ContentReportBuilder.OpenXml.Builder.Word.Elements;
using System.Collections.Generic;


namespace ContentReportBuilder.OpenXml.Builder
{
    public class DocumentElementFactory: DocumentElementFactoryBase
    {
        protected override Dictionary<DocumentElementType, IDocumentElementBuilder> ExcelElements()
        {
            var elements = new Dictionary<DocumentElementType, IDocumentElementBuilder>();
            elements.Add(DocumentElementType.Heading, new ExcelHeading());
            elements.Add(DocumentElementType.Paragraph, new ExcelParagraph());
            elements.Add(DocumentElementType.List, new ExcelList());
            elements.Add(DocumentElementType.Table, new ExcelTable());
            elements.Add(DocumentElementType.LineChart, new ExcelLineChart());
            return elements;
        }

        protected override Dictionary<DocumentElementType, IDocumentElementBuilder> WordElements()
        {
            var elements = new Dictionary<DocumentElementType, IDocumentElementBuilder>();
             elements.Add(DocumentElementType.Heading, new WordHeading());
            elements.Add(DocumentElementType.Paragraph, new WordParagraph());
            elements.Add(DocumentElementType.List, new WordList());
            elements.Add(DocumentElementType.Table, new WordTable());
            elements.Add(DocumentElementType.Image, new WordImage());
            return elements;
        }
    }
}
