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
    public class PdfPieChart : DocumentElementBuilder<Document>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.PieChart;

        protected override void Build(DocumentModelElement element, Document documentSection)
        {
            ChartData elementData = (ChartData)element.GetData();

            var plotModel = new PlotModel { Title = elementData.Title };
            var pieSeries = new OxyPlot.Series.PieSeries();

            foreach(var item in elementData.Series ?? new List<ChartSeries>())
            {
                pieSeries.Slices.Add(new PieSlice(item.Title, item.BarValue) { IsExploded = item.IsExploded });
            }
            
            pieSeries.InnerDiameter = 0;
            pieSeries.ExplodedDistance = 0.0;
            pieSeries.Stroke = OxyColors.White;
            pieSeries.StrokeThickness = 2.0;
            pieSeries.InsideLabelPosition = 0.8;
            pieSeries.AngleSpan = 360;
            pieSeries.StartAngle = 0;
            plotModel.Series.Add(pieSeries);

            WritePlotToDocument(documentSection, plotModel, 240, 160);
        }
    }
}
