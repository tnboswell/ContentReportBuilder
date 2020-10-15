using ContentReportBuilder.Builder;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Reflection;

namespace ContentReportBuilder.OpenXml.Builder
{
    public abstract class TableElementBuilderBase<T> : DocumentElementBuilderBase<T>, IDocumentElementBuilder
    {

        protected List<string> Headers(object row)
        {
            List<string> headers = new List<string>();
            var objectType = row.GetType();
            foreach (var property in objectType.GetProperties(BindingFlags.Public))
            {
                headers.Add(property.Name);
            }
            return headers;
        }

        protected string CellValue(object row, string key)
        {
            var objectType = row.GetType();

            PropertyInfo numberPropertyInfo = objectType.GetProperty(key);

            // get value of property: public double Number
            return numberPropertyInfo.GetValue(row).ToString();
        }

        protected Cell CreateCell(string text)
        {
            Cell cell = new Cell();
            cell.StyleIndex = 1U;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);
            return cell;
        }
        protected Cell CreateCell(string text, uint styleIndex)
        {
            Cell cell = new Cell();
            cell.StyleIndex = styleIndex;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);
            return cell;
        }

        protected EnumValue<CellValues> ResolveCellDataTypeOnValue(string text)
        {
            int intVal;
            double doubleVal;
            if (int.TryParse(text, out intVal) || double.TryParse(text, out doubleVal))
            {
                return CellValues.Number;
            }
            else
            {
                return CellValues.String;
            }
        }
    }


}
