using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using iTextSharp.text;
using System;

namespace ContentReportBuilder.ITextSharp.Builder.Pdf.Sections
{
    public class PdfSectionBuilder : DocumentSectionBuilderBase<Document>, IDocumentSectionBuilder<Document>
    {
        public PdfSectionBuilder():base(new DocumentElementFactory())
        {

        }
    }
}
