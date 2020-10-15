using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace ContentReportBuilder.docX.Builder.Word.Elements
{
    public class WordLineChart : DocumentElementBuilderBase<DocX>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.LineChart;

        protected override void Build(DocumentModelElement element, DocX documentSection)
        {
            ChartData elementData = (ChartData)element.GetData();

            // Create a line chart.
            var lineChart = new LineChart();
            lineChart.AddLegend(ChartLegendPosition.Left, false);
           
            foreach (var value in elementData.Series ?? new List<ChartSeries>())
            {
                var series = new Series(value.Title);
                series.Bind(value.Items.Select(a => a.X).ToList(), value.Items.Select(a => a.Y).ToList());
                lineChart.AddSeries(series);
            }

            documentSection.InsertChart(lineChart);

        }
    }
}
