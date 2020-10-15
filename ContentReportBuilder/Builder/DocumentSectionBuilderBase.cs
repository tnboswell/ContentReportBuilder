using ContentReportBuilder.Models;
using System.Collections.Generic;

namespace ContentReportBuilder.Builder
{
    public abstract class DocumentSectionBuilderBase<T>: IDocumentSectionBuilder<T>
    {
        public DocumentSectionBuilderBase(DocumentElementFactoryBase documentFactory)
        {
            ElementFactory = documentFactory;
        }
        protected DocumentElementFactoryBase ElementFactory { get; set; }
        public void Build(DocumentBuilderType type, DocumentModelSection section, T document)
        {
            foreach (var element in section.Elements ?? new List<DocumentModelElement>())
            {
                var elementBuilder = GetElementBuilder(type, element.Type);
                if (elementBuilder == null)
                {
                    // some log msg to say builder skipped
                }
                elementBuilder.Build(element, document);
            }
        }

        public virtual IDocumentElementBuilder GetElementBuilder(DocumentBuilderType type, DocumentElementType elementType)
        {

            return ElementFactory.GetElementBuilder(type, elementType);
        }

    }
}
