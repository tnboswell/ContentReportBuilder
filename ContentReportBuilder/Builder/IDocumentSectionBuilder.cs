using ContentReportBuilder.Models;

namespace ContentReportBuilder.Builder
{
    public interface IDocumentSectionBuilder<T>
    {
        void Build(DocumentBuilderType type, DocumentModelSection section, T document);
        // will be used to check the dynamic type is correct, for instance a word doc or excel doc
    }
}
