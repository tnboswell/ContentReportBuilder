using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.Styles;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;

namespace ContentReportBuilder.EPPlus.Builder.Excel.Elements
{
    public class ExcelBarChart : DocumentElementBuilder<ExcelWorksheet>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.BarChart;

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
                // series Header
                documentSection.Cells[row, col].Value = value.Title;
                documentSection.Cells[row + 1, col].Value = value.BarValue;
                col++;
            }

            col--;

            var barChart = documentSection.Drawings.AddChart("chart", OfficeOpenXml.Drawing.Chart.eChartType.ColumnStacked);
            barChart.SetPosition(rowStart - 1, 0, colStart - 1, 0);
            barChart.SetSize(400, 400);

            var headerRange = ExcelRange.GetAddress(rowStart, colStart, rowStart, col);
            var valuesRange = ExcelRange.GetAddress(rowStart + 1, colStart, rowStart + 1, col);
            var series = barChart.Series.Add(valuesRange, headerRange);

            barChart.Legend.Position = OfficeOpenXml.Drawing.Chart.eLegendPosition.Right;
        }
    }
}
