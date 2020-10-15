using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ContentReportBuilder.OpenXml.Builder.Excel.Elements
{
    public class ExcelLineChart : DocumentElementBuilderBase<Sheet>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.LineChart;

        protected override void Build(DocumentModelElement element, Sheet documentSection)
        {
            ChartData elementData = (ChartData)element.GetData();
            throw new NotImplementedException();
        }
    }
}
