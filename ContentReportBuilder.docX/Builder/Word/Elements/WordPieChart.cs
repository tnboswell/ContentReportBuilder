using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace ContentReportBuilder.docX.Builder.Word.Elements
{
    public class WordPieChart : DocumentElementBuilderBase<DocX>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.PieChart;

        protected override void Build(DocumentModelElement element, DocX documentSection)
        {
            ChartData elementData = (ChartData)element.GetData();
            // Create a pie chart.
            var pieChart = new PieChart();
            pieChart.AddLegend(ChartLegendPosition.Left, false);

            var series = new Series("");
            series.Bind(elementData.Series.Select(a => a.Title).ToList(), elementData.Series.Select(a => a.BarValue).ToList());
            pieChart.AddSeries(series);
            documentSection.InsertChart(pieChart);

        }
    }
}
