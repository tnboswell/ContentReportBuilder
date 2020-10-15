using DocumentFormat.OpenXml.Wordprocessing;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using System.Collections.Generic;
using ContentReportBuilder.Builder;

//http://www.ludovicperrichon.com/create-a-word-document-with-openxml-and-c/#createword
namespace ContentReportBuilder.OpenXml.Builder.Word.Elements
{
    public class WordTable : TableElementBuilderBase<Body>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Table;

        protected override void Build(DocumentModelElement element, Body documentSection)
        {
            TableData elementData = (TableData)element.GetData();

            Table table = new Table();
            List<string> headers = null;
            if (elementData != null && elementData.Rows != null && elementData.Rows.Count > 0)
            {
                headers = Headers(elementData.Rows);
                TableRow tableHeader = new TableRow();
                foreach (var headerItem in headers)
                {
                    TableHeader headerCell = new TableHeader();
                    Paragraph paragraph = new Paragraph(new Run(new Text(headerItem)));
                    headerCell.Append(paragraph);
                    tableHeader.Append(headerCell);
                }
                


                if (headers.Count == 0)
                    return;

                foreach (var row in elementData.Rows)
                {
                    TableRow tableRow = new TableRow();
                    
                    foreach (var headerCell in headers)
                    {
                        TableCell tableCell = new TableCell();
                        Paragraph paragraph = new Paragraph(new Run(new Text(CellValue(row, headerCell).ToString())));
                        tableCell.Append(paragraph);
                        tableRow.Append(tableCell);

                    }

                }
            }

        }
    }
}
