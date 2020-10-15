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
    public class ExcelPieChart : DocumentElementBuilder<ExcelWorksheet>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.PieChart;

        protected override void Build(DocumentModelElement element, ExcelWorksheet documentSection)
        {
            ChartData elementData = (ChartData)element.GetData();
            IExcelElementStyle style = element.ExcelStyle;

            int row = style.Cell.NextRow(NextRow(documentSection));
            int col = style.Cell.NextColumn(NextColumn(documentSection, row));

            int colStart = col;
            foreach (var item in elementData.Series ?? new List<ChartSeries>())
            {

                documentSection.Cells[row, col].Value = item.Title; // header
                documentSection.Cells[(row + 1), col].Value = item.BarValue; // value
                col++;
            }

            col--;

            //Add the piechart
            var pieChart = documentSection.Drawings.AddChart("crtExtensionsSize", OfficeOpenXml.Drawing.Chart.eChartType.PieExploded3D);
            //Set top left corner to row 1 column 2
            pieChart.SetPosition(row, 0, colStart, 0);
            pieChart.SetSize(400, 400);
            var headerRange = ExcelRange.GetAddress(row, colStart, row, col);
            var valuesRange = ExcelRange.GetAddress((row + 1), colStart, (row + 1), col);
            pieChart.Series.Add(valuesRange, headerRange);

            pieChart.Title.Text = "Extension Size";
            //Set datalabels and remove the legend
            // pieChart.ShowCategory = true;
            // pieChart.DataLabel.ShowPercent = true;
            // pieChart.DataLabel.ShowLeaderLines = true;
            pieChart.Legend.Position = OfficeOpenXml.Drawing.Chart.eLegendPosition.Bottom;
        }
    }
}
