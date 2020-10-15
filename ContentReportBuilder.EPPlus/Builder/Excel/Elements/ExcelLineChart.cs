using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.Styles;
using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace ContentReportBuilder.EPPlus.Builder.Excel.Elements
{
    public class ExcelLineChart : DocumentElementBuilder<ExcelWorksheet>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.LineChart;

        protected override void Build(DocumentModelElement element, ExcelWorksheet documentSection)
        {
            ChartData elementData = (ChartData)element.GetData();
            IExcelElementStyle style = element.ExcelStyle;

            int row = style.Cell.NextRow(NextRow(documentSection));
            int col = style.Cell.NextColumn(NextColumn(documentSection, row));
            int rowStart = row;
            int colStart = col;
            foreach (var value in elementData.Series ?? new List<ChartSeries>())
            {
                col = colStart;
                // series Header
                documentSection.Cells[(row + 1), col].Value = value.Title;
                col++;
                foreach (var item in value.Items ?? new List<SeriesItem>())
                {
                    documentSection.Cells[row, col].Value = item.X;
                    documentSection.Cells[(row + 1), col].Value = item.Y;
                    col++;
                }
                row += 2;

            }

            col--;

            //create a new piechart of type Line
            var lineChart = documentSection.Drawings.AddChart("lineChart", OfficeOpenXml.Drawing.Chart.eChartType.Line);
            lineChart.SetPosition(rowStart, 0, colStart, 0);
            lineChart.SetSize(400, 400);

            for(int i = rowStart; i < row; i += 2)
            {
                var headerRange = ExcelRange.GetAddress(i, colStart + 1, i, col);
                var valuesRange = ExcelRange.GetAddress(i + 1, colStart + 1, i + 1, col);
                var series = lineChart.Series.Add(valuesRange, headerRange);
                series.Header = documentSection.Cells[i + 1, colStart].Value.ToString();
            }

            lineChart.Legend.Position = OfficeOpenXml.Drawing.Chart.eLegendPosition.Right;
        }
    }
}
