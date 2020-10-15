using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using iTextSharp.text;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using System;
using System.Collections.Generic;
using System.IO;

namespace ContentReportBuilder.ITextSharp.Builder.Pdf.Elements
{
    public class PdfLineChart : DocumentElementBuilder<Document>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.LineChart;

        protected override void Build(DocumentModelElement element, Document documentSection)
        {
            ChartData elementData = (ChartData)element.GetData();

            var plotModel = new PlotModel { Title = elementData.Title };
            plotModel.Axes.Add(new OxyPlot.Axes.LinearAxis { Position = AxisPosition.Left, Title = "Y-Axis" });
            plotModel.Axes.Add(new OxyPlot.Axes.LinearAxis { Position = AxisPosition.Bottom, Title = "X-Axis" });

            foreach (var value in elementData.Series ?? new List<ChartSeries>())
            {
                var series = new OxyPlot.Series.LineSeries()
                {
                    Title = value.Title,
                    Color = OxyColors.SkyBlue,
                    MarkerType = MarkerType.Circle,
                    MarkerSize = 6,
                    MarkerStroke = OxyColors.White,
                    MarkerFill = OxyColors.SkyBlue,
                    MarkerStrokeThickness = 1.5
                };
                foreach(var item in value.Items ?? new List<SeriesItem>())
                {
                    series.Points.Add(new DataPoint(item.X, item.Y));
                }
                plotModel.Series.Add(series);
            }

            WritePlotToDocument(documentSection, plotModel, 240, 160);
        }
    }
}
