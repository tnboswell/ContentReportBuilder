using ContentReportBuilder.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ContentReportBuilder
{
    /// <summary>
    /// Interface to create build a document using a provided document model for a requested document type
    /// </summary>
    public interface IDocumentDirector
    {
        /// <summary>
        /// Builds a Stream containing the document model in the format in the requested document type.
        /// The Stream is to be returned via a HTTP response
        /// </summary>
        /// <param name="type"></param>
        /// <param name="document"></param>
        /// <returns>A stream for returning as a response to a HTTP request for the document</returns>
        Stream BuildStream(DocumentBuilderType type, DocumentModel document, out string errorMsg);
        /// <summary>
        /// Builds and saves a document to file for a provided document model for a requested document type
        /// </summary>
        /// <param name="type"></param>
        /// <param name="document"></param>
        /// <returns>indicates file created and saved succcessfully</returns>
        bool BuildFile(DocumentBuilderType type, DocumentModel document, out string errorMsg);

    }
}
