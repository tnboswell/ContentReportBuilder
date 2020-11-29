using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.ElementModels;
using ContentReportBuilder.Models.Styles;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ContentReportBuilder.EPPlus.Builder.Excel.Elements
{
    public class ExcelTable : DocumentElementBuilder<ExcelWorksheet>, IDocumentElementBuilder
    {
        protected readonly TableDataAdapter _adapter = new TableDataAdapter();
        public override DocumentElementType Type => DocumentElementType.Table;

        protected override void Build(DocumentModelElement element, ExcelWorksheet documentSection)
        {
            TableData elementData = (TableData)element.GetData();
            IExcelElementStyle style = element.ExcelStyle;

            var colStyles = ((TableElementModel)element).ColStyle;
            var headingColour = ((TableElementModel)element).HeadingColour;

            if (elementData.Rows == null || elementData.Rows.Count() == 0)
                return;

            var headers = _adapter.Headers(elementData.Rows.First());

            int rowStart = style.Cell.NextRow(NextRow(documentSection));
            int colStart = style.Cell.NextColumn(NextColumn(documentSection, rowStart));

            var indexes = (headers.Count);
            var rows = (elementData.Rows.Count * nextRow(style, 0) + 1);
            int row = rowStart;
            int col = colStart;
            int index = 0;

            while(index < indexes)
            {
                var currentStyle = style;
                var styleIndex = index + 1;
                if (colStyles != null && colStyles.ContainsKey(styleIndex))
                    currentStyle = colStyles[styleIndex];

                CellStyle(documentSection, row, col, currentStyle, headingColour);
                
                documentSection.Cells[row, col].Value = ToDataType(currentStyle.Cell.DataType, _adapter.ReadableHeader(headers[index]));
                index++;
                col = nextCol(currentStyle, col);
            }
            row = nextRow(style, row);
            int headerIndex = 0;
            int rowIndex = 0;
            for (; row < (rows + rowStart);)
            {
                headerIndex = 0;
                //for (col = colStart; col < (cols  + colStart);)
                col = colStart;
                while (headerIndex < indexes)
                {
                    var currentStyle = style;
                    var styleIndex = headerIndex + 1;
                    if (colStyles != null && colStyles.ContainsKey(styleIndex))
                        currentStyle = colStyles[styleIndex];

                    CellStyle(documentSection, row, col, currentStyle);
                    documentSection.Cells[row, col].Value = ToDataType(currentStyle.Cell.DataType, _adapter.CellValue(elementData.Rows[rowIndex], headers[headerIndex]));
                    headerIndex++;
                    col = nextCol(currentStyle, col);
                }
                rowIndex++;
                row = nextRow(style, row);
            }
            col--;
            row--;

            if (addTable(style))
            {
                //create a range for the table
                ExcelRange tableRange = documentSection.Cells[rowStart, colStart, row, col];
                //add a table to the range
                var tableName = $"Table{documentSection.Name}{documentSection.Tables.Count()}";
                var tab = documentSection.Tables.Add(tableRange, tableName.Replace(" ", ""));
            }
        }

        private bool addTable(IExcelElementStyle style)
        {
            return (style == null || style.Cell.CellMerging == null);
        }

        private int nextCol(IExcelElementStyle style, int col)
        {
            if (style == null || style.Cell.CellMerging == null)
                return col + 1;
            else
                return style.Cell.CellMerging.ColumnsMerge(col) + 1;
        }

        private int nextRow(IExcelElementStyle style, int row)
        {
            if (style == null || style.Cell.CellMerging == null)
                return row + 1;
            else
                return style.Cell.CellMerging.RowsMerge(row) + 1;
        }
    }
}
