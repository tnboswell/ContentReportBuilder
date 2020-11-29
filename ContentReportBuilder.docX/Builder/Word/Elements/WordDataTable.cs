using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace ContentReportBuilder.docX.Builder.Word.Elements
{
    public class WordDataTable : DocumentElementBuilderBase<DocX>, IDocumentElementBuilder
    {
        protected readonly TableDataAdapter _adapter = new TableDataAdapter();
        public override DocumentElementType Type => DocumentElementType.Table;

        protected override void Build(DocumentModelElement element, DocX documentSection)
        {
            TableData elementData = (TableData)element.GetData();

            if (elementData.Rows == null || elementData.Rows.Count() == 0)
                return;

            var headers = _adapter.Headers(elementData.Rows.First());

            // Add a Table into the document and sets its values.
            var table = documentSection.AddTable((elementData.Rows.Count + 1), (headers.Count));
            table.Design = TableDesign.ColorfulListAccent1;
            table.Alignment = Alignment.center;
            var cols = (headers.Count);
            var rows = (elementData.Rows.Count + 1);

            int row = 0;
            for(int col = 0; col < cols; col++)
            {
                table.Rows[row].Cells[col].Paragraphs[0].Append(_adapter.ReadableHeader(headers[col]));
            }

            for(row = 1; row < rows;  row++ )
            {
                for (int col = 0; col < cols; col++)
                {
                    table.Rows[row].Cells[col].Paragraphs[0].Append(_adapter.CellValue(elementData.Rows[(row - 1)], headers[col]));
                }
            }
            //table.SetBorder(TableBorderType.Bottom,);
            documentSection.InsertTable(table);
        }

    }
}
