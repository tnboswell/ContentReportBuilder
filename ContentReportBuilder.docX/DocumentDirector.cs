using ContentReportBuilder.Builder;
using ContentReportBuilder.docX.Builder.Word;
using ContentReportBuilder.Models;
using System;
using System.IO;

namespace ContentReportBuilder.docX
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
            if (type == DocumentBuilderType.Word)
            {
                return new WordDocumentBuilder();
            }

            return null;
        }
    }
}
