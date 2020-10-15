using ContentReportBuilder.Builder;
using ContentReportBuilder.Models.Styles;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Drawing;

namespace ContentReportBuilder.EPPlus.Builder
{
    public abstract class DocumentElementBuilder<T> : DocumentElementBuilderBase<T>, IDocumentElementBuilder
    {

        protected void CellStyle(ExcelWorksheet documentSection, int row, int col, IExcelElementStyle style, Color? overrideColour = null)
        {
            if (style == null)
                return;

            int colWidth = col;
            int rowWidth = row;
            if (style.Cell.CellMerging == null)
            {
                documentSection.Row(row).Height = style.Cell.Height;
                documentSection.Column(col).Width = style.Cell.Width;
            }
            else
            {
                colWidth = style.Cell.CellMerging.ColumnsMerge(col);
                rowWidth = style.Cell.CellMerging.RowsMerge(row);
                documentSection.Cells[row, col, rowWidth, colWidth].Merge = true; //[FromRow, FromCol, ToRow, ToCol]
            }

            documentSection.Cells[row, col, rowWidth, colWidth].Style.Font.Bold = style.Font.IsBold;
            documentSection.Cells[row, col, rowWidth, colWidth].Style.Font.Italic = style.Font.IsItalics;
            documentSection.Cells[row, col, rowWidth, colWidth].Style.Font.Strike = style.Font.IsStrikeThrough;
            documentSection.Cells[row, col, rowWidth, colWidth].Style.Font.UnderLine = style.Font.IsUnderlined;
            documentSection.Cells[row, col, rowWidth, colWidth].Style.Font.Name = style.Font.Font;
            documentSection.Cells[row, col, rowWidth, colWidth].Style.Font.Size = style.Font.Size;
            documentSection.Cells[row, col, rowWidth, colWidth].Style.WrapText = style.Cell.IsWrapped;
            documentSection.Cells[row, col, rowWidth, colWidth].Style.Font.Color.SetColor(style.Font.Colour);
            documentSection.Cells[row, col, rowWidth, colWidth].Style.Fill.PatternType = (style.Cell.FillSolid) ? ExcelFillStyle.Solid : ExcelFillStyle.None;


            if(overrideColour != null)
            {
                documentSection.Cells[row, col, rowWidth, colWidth].Style.Fill.PatternType = ExcelFillStyle.Solid;
                documentSection.Cells[row, col, rowWidth, colWidth].Style.Fill.BackgroundColor.SetColor(overrideColour.Value);
            }
            else if (style.Cell.FillSolid)
            {
                documentSection.Cells[row, col, rowWidth, colWidth].Style.Fill.BackgroundColor.SetColor(style.Cell.Colour);
            }

            documentSection.Cells[row, col, rowWidth, colWidth].Style.HorizontalAlignment = GetHorizontalAlignment(style.HorizontalAlignment);
            documentSection.Cells[row, col, rowWidth, colWidth].Style.VerticalAlignment = GetVerticalAlignment(style.VeritcalAlignment);

            documentSection.Cells[row, col, rowWidth, colWidth].Style.Border.Top.Style = BorderStyle(style.Cell.Border.Top);
            documentSection.Cells[row, col, rowWidth, colWidth].Style.Border.Bottom.Style = BorderStyle(style.Cell.Border.Bottom);
            documentSection.Cells[row, col, rowWidth, colWidth].Style.Border.Left.Style = BorderStyle(style.Cell.Border.Left);
            documentSection.Cells[row, col, rowWidth, colWidth].Style.Border.Right.Style = BorderStyle(style.Cell.Border.Right);
        }

        protected object ToDataType(CellDataType dataType, string value)
        {
            switch (dataType)
            {
                case CellDataType.String:
                    return value;
                case CellDataType.Integer:
                    int intVal;
                    if(int.TryParse(value, out intVal))
                    {
                        return intVal;
                    }
                    return value;
                case CellDataType.Double:
                    double doubleVal;
                    if (double.TryParse(value, out doubleVal))
                    {
                        return doubleVal;
                    }
                    return value;
                case CellDataType.Date:
                    DateTime dateVal;
                    if (DateTime.TryParse(value, out dateVal))
                    {
                        return dateVal;
                    }
                    return value;
                default:
                    return value;
            }
        }

        protected ExcelBorderStyle BorderStyle(CellBorder border)
        {
            switch (border)
            {
                case CellBorder.None:
                    return ExcelBorderStyle.None;
                case CellBorder.Thin:
                    return ExcelBorderStyle.Thin;
                case CellBorder.Dashed:
                    return ExcelBorderStyle.Dashed;
                case CellBorder.Dotted:
                    return ExcelBorderStyle.Dotted;
                case CellBorder.Medium:
                    return ExcelBorderStyle.Medium;
                case CellBorder.Thick:
                    return ExcelBorderStyle.Thick;
                default:
                    return ExcelBorderStyle.None;
            }
        }
        protected ExcelHorizontalAlignment GetHorizontalAlignment(HorizontalAlignment alignment)
        {
            switch (alignment)
            {
                case HorizontalAlignment.Centre:
                    return ExcelHorizontalAlignment.Center;
                case HorizontalAlignment.Justified:
                    return ExcelHorizontalAlignment.Justify;
                case HorizontalAlignment.Left:
                    return ExcelHorizontalAlignment.Left;
                case HorizontalAlignment.Right:
                    return ExcelHorizontalAlignment.Right;
                default:
                    return ExcelHorizontalAlignment.Center;
            }
        }

        protected ExcelVerticalAlignment GetVerticalAlignment(VeritcalAlignment alignment)
        {
            switch (alignment)
            {
                case VeritcalAlignment.Middle:
                    return ExcelVerticalAlignment.Center;
                case VeritcalAlignment.Bottom:
                    return ExcelVerticalAlignment.Bottom;
                case VeritcalAlignment.Top:
                    return ExcelVerticalAlignment.Top;
                default:
                    return ExcelVerticalAlignment.Center;
            }
        }

        protected int NextRow(ExcelWorksheet documentSection)
        {
            var dimensions = documentSection.Dimension;

            // starts at zero as this will result in a 1 offset equaling row 1
            if (dimensions == null)
                return 0;

            return documentSection.Dimension.Rows;
        }
        protected int NextColumn(ExcelWorksheet documentSection, int row)
        {
            var dimensions = documentSection.Dimension;
            if (dimensions == null)
                return 1;

            int maxWidth = documentSection.Dimension.Columns;
            int startCol = 1;
            for (int col = startCol; col <= maxWidth; col++)
            {
                if (documentSection.Cells[row, col].Value == null &&
                    documentSection.Cells[row, col].Merge == false)
                {
                    return col - 1;
                }
            }
            return maxWidth - 1;
        }
    }


}
