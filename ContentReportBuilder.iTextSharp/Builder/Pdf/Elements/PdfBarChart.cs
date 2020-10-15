using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using iTextSharp.text;
using OxyPlot;
using OxyPlot.Series;
using System.Collections.Generic;


namespace ContentReportBuilder.ITextSharp.Builder.Pdf.Elements
{
    public class PdfBarChart : DocumentElementBuilder<Document>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.BarChart;

        protected override void Build(DocumentModelElement element, Document documentSection)
        {
            ChartData elementData = (ChartData)element.GetData();
            var plotModel = new PlotModel { Title = elementData.Title };

            var barSeries = new OxyPlot.Series.BarSeries
            { LabelPlacement = LabelPlacement.Inside };
            var itemsSource = new List<BarItem>();
            foreach (var value in elementData.Series ?? new List<ChartSeries>())
            {
                itemsSource.Add(new BarItem(value.BarValue));
                  
            }
            barSeries.ItemsSource = itemsSource;
            plotModel.Series.Add(barSeries);

            WritePlotToDocument(documentSection, plotModel, 240, 160);
            
        }
    }
}
