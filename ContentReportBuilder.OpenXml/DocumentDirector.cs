using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.OpenXml.Builder.Excel;
using ContentReportBuilder.OpenXml.Builder.Word;
using System.IO;
using System;

namespace ContentReportBuilder.OpenXml
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
            throw new NotImplementedException("OpenXml is not currently supported, please use docx and EPPlus libaries.");
            if (type == DocumentBuilderType.Excel)
            {
                return new ExcelDocumentBuilder();
            }
            else if (type == DocumentBuilderType.Word)
            {
                return new WordDocumentBuilder();
            }

            return null;
        }
    }
}
