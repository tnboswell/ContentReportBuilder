using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

//http://www.dispatchertimer.com/tutorial/how-to-create-an-excel-file-in-net-using-openxml-part-1-basics/

// Excel table example
// https://www.c-sharpcorner.com/article/creating-excel-file-using-openxml/
namespace ContentReportBuilder.OpenXml.Builder.Excel.Elements
{
    public class ExcelTable : TableElementBuilderBase<Sheet>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Table;

        protected override void Build(DocumentModelElement element, Sheet documentSection)
        {
            TableData elementData = (TableData)element.GetData();
            SheetData tableData = new SheetData();
            List<string> headers = null;
            if (elementData != null && elementData.Rows != null && elementData.Rows.Count > 0)
            {
                headers = Headers(elementData.Rows);

                if (headers.Count == 0)
                    return;


                Row headerRow = new Row();
                foreach (var headerItem in headers)
                {
                    headerRow.Append(CreateCell(headerItem, 2U));
                }
                tableData.Append(headerRow);
                
                foreach(var row in elementData.Rows)
                {
                    Row tableRow = new Row();
                    foreach (var header in headers)
                    {
                        tableRow.Append(CreateCell(CellValue(row, header)));
                    }
                    tableData.Append(tableRow);
                }
                documentSection.Append(tableData);
            }
           
        }

    }
}
