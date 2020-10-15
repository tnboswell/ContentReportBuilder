using ContentReportBuilder.Builder;
using ContentReportBuilder.ITextSharp.Builder.Pdf;
using ContentReportBuilder.Models;
using System.IO;

namespace ContentReportBuilder.ITextSharp
{
    public class DocumentDirector : DocumentDirectorBase, IDocumentDirector
    {

        /// <summary>
        /// Gets a requested document builder type
        /// </summary>
        /// <param name="type"></param>
        /// <returns>returns the requested document builder</returns>
        protected override IDocumentBuilder GetBuilder(DocumentBuilderType type)
        {
            if (type == DocumentBuilderType.Pdf)
            {
                return new PdfDocumentBuilder();
            }

            return null;
        }

    }
}
