using ContentReportBuilder.Builder;
using ContentReportBuilder.EPPlus.Builder.Excel.Elements;
using ContentReportBuilder.Models;
using System;
using System.Collections.Generic;


namespace ContentReportBuilder.EPPlus.Builder
{
    public class DocumentElementFactory: DocumentElementFactoryBase
    {
        protected override Dictionary<DocumentElementType, IDocumentElementBuilder> ExcelElements()
        {
            var elements = new Dictionary<DocumentElementType, IDocumentElementBuilder>();
            elements.Add(DocumentElementType.BarChart, new ExcelBarChart());
            elements.Add(DocumentElementType.Heading, new ExcelHeading());
            elements.Add(DocumentElementType.Image, new ExcelImage());
            elements.Add(DocumentElementType.LineChart, new ExcelLineChart());
            elements.Add(DocumentElementType.List, new ExcelList());
            elements.Add(DocumentElementType.Paragraph, new ExcelParagraph());
            elements.Add(DocumentElementType.PieChart, new ExcelPieChart()) ;
            elements.Add(DocumentElementType.Table, new ExcelTable());
            elements.Add(DocumentElementType.Header, new ExcelHeader());
            elements.Add(DocumentElementType.Footer, new ExcelFooter());

            return elements;
        }


    }

}
