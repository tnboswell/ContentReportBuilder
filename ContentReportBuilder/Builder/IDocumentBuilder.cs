using ContentReportBuilder.Models;
using System.IO;


namespace ContentReportBuilder.Builder
{
    public interface IDocumentBuilder
    {
        DocumentBuilderType Type { get; }
        Stream BuildStream(DocumentModel document);
        void BuildFile(DocumentModel document);

    }
}
