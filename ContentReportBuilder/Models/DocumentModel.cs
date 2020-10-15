using ContentReportBuilder.Models.MetaData;
using ContentReportBuilder.Models.Styles;
using System;
using System.Collections.Generic;

namespace ContentReportBuilder.Models
{
    public class DocumentModel
    {
        public DocumentModel(IDocumentStyle style, IDocumentMetaData metaData)
        {
            MetaData = metaData ?? throw new Exception("Meta Data provided to Document Model is empty;");
            Style = style;
            DocumentSections = new List<DocumentModelSection>();
        }
        public IDocumentMetaData MetaData { get; set; }
        public IDocumentStyle Style { get; set; }

        // look at using a document section strategy
        public List<DocumentModelSection> DocumentSections { get; set; }

        public void AddDocumentSection(DocumentModelSection section)
        {
            DocumentSections.Add(section);
        }
    }
}
