using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ContentReportBuilder.ITextSharp.Builder.Pdf.Elements
{
    public class PdfDataTable : DocumentElementBuilderBase<Document>, IDocumentElementBuilder
    {
        protected readonly TableDataAdapter _adapter = new TableDataAdapter();
        public override DocumentElementType Type => DocumentElementType.Table;
        //https://www.sharepointpals.com/post/generate-pdf-report-using-itextsharp-net-pdf-library-in-sharepoint-environment/
        protected override void Build(DocumentModelElement element, Document documentSection)
        {
            TableData elementData = (TableData)element.GetData();
            if (elementData.Rows == null || elementData.Rows.Count() == 0)
                return;
            var headers = _adapter.Headers(elementData.Rows.First());
            var cols = (headers.Count);
            var rows = (elementData.Rows.Count + 1);

            var table = new PdfPTable(cols);

            int row = 0;
            for (int col = 0; col < cols; col++)
            {
                table.AddCell(_adapter.ReadableHeader(headers[col]));
            }

            for (row = 1; row < rows; row++)
            {
                for (int col = 0; col < cols; col++)
                {
                    table.AddCell(_adapter.CellValue(elementData.Rows[(row - 1)], headers[col]));
                }
            }
            
            documentSection.Add(table);
        }


    }
}
