using ContentReportBuilder.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace ContentReportBuilder.Builder
{
    public abstract class DocumentElementFactoryBase
    {
        public DocumentElementFactoryBase()
        {
            RegisteredElementBuilders = new Dictionary<DocumentBuilderType, Dictionary<DocumentElementType, IDocumentElementBuilder>>();
            RegisteredElementBuilders.Add(DocumentBuilderType.Excel, ExcelElements());
            RegisteredElementBuilders.Add(DocumentBuilderType.Word, WordElements());
            RegisteredElementBuilders.Add(DocumentBuilderType.Pdf, PdfElements());

        }

        protected Dictionary<DocumentBuilderType, Dictionary<DocumentElementType, IDocumentElementBuilder>> RegisteredElementBuilders { get; set; }
        public IDocumentElementBuilder GetElementBuilder(DocumentBuilderType documentType, DocumentElementType elementType)
        {
            if (RegisteredElementBuilders.ContainsKey(documentType) && RegisteredElementBuilders[documentType].ContainsKey(elementType))
            {
                return RegisteredElementBuilders[documentType][elementType];
            }

            return null;
        }

        protected virtual Dictionary<DocumentElementType, IDocumentElementBuilder> ExcelElements()
        {
            return new Dictionary<DocumentElementType, IDocumentElementBuilder>();
        }

        protected virtual Dictionary<DocumentElementType, IDocumentElementBuilder> WordElements()
        {
            return new Dictionary<DocumentElementType, IDocumentElementBuilder>();
        }

        protected virtual Dictionary<DocumentElementType, IDocumentElementBuilder> PdfElements()
        {
            return new Dictionary<DocumentElementType, IDocumentElementBuilder>();
        }
    }
}
