using ContentReportBuilder.Models;

namespace ContentReportBuilder.Builder
{
    public interface IDocumentElementBuilder
    {
        DocumentElementType Type { get; }
        void Build(dynamic element, dynamic documentSection);

    }
}
