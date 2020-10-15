using ContentReportBuilder.Models;

namespace ContentReportBuilder.Builder
{
    public abstract class DocumentElementBuilderBase<T> : IDocumentElementBuilder
    {
        public abstract DocumentElementType Type { get; }

        protected abstract void Build(DocumentModelElement element, T documentSection);

        public void Build(dynamic element, dynamic documentSection)
        {
            Build(element, documentSection);
        }
    }
}
