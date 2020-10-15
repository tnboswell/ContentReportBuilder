using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using System.Linq;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace ContentReportBuilder.docX.Builder.Word.Elements
{
    public class WordBarChart : DocumentElementBuilderBase<DocX>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.BarChart;

        protected override void Build(DocumentModelElement element, DocX documentSection)
        {
            ChartData elementData = (ChartData)element.GetData();
            // Create a bar chart.
            var barChart = new BarChart();
            barChart.AddLegend(ChartLegendPosition.Left, false);
            barChart.BarDirection = BarDirection.Bar;
            barChart.BarGrouping = BarGrouping.Standard;
            barChart.GapWidth = 200;

            // Create and add series
            var series = new Series("");
            series.Bind(elementData.Series.Select(a => a.Title).ToList(), elementData.Series.Select(a => a.BarValue).ToList());
            barChart.AddSeries(series);
            documentSection.InsertChart(barChart, 350, 550);
        }
    }
}
