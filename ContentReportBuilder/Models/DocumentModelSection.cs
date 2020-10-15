using ContentReportBuilder.Models.MetaData;
using System.Collections.Generic;
using System.Linq;

namespace ContentReportBuilder.Models
{
    /// <summary>
    /// Used to define either a Word Section or Excel Sheet
    /// </summary>
    public class DocumentModelSection
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="metaData">Document Section meta data must be provided as it is used to build the document section</param>
        public DocumentModelSection(ISectionMetaData metaData)
        {
            MetaData = metaData;
            Elements = new List<DocumentModelElement>();
        }

        public DocumentModelSection(ISectionMetaData metaData, List<DocumentModelElement> elements)
        {
            MetaData = metaData;
            Elements = new List<DocumentModelElement>();
            AddElementRange(elements);
        }

        public void AddElement(DocumentModelElement element)
        {
            Elements.Add(element);
        }

        public void AddElementRange(List<DocumentModelElement> elements)
        {
            foreach(var element in elements)
                AddElement(element);
        }

        // look at using a document element strategy
        public ISectionMetaData MetaData { get; set; }
        public List<DocumentModelElement> Elements { get; set; }

    }
}
