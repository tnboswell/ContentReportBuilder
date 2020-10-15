using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ContentReportBuilder
{
    public abstract class DocumentDirectorBase : IDocumentDirector
    {
        /// <summary>
        /// Builds and saves a document to file for a provided document model for a requested document type
        /// </summary>
        /// <param name="type"></param>
        /// <param name="document"></param>
        /// <returns>indicates file created and saved succcessfully</returns>
        public bool BuildFile(DocumentBuilderType type, DocumentModel document, out string errorMsg)
        {
            errorMsg = "";
            var builder = GetBuilder(type);

            if (builder == null)
            {
                errorMsg = $"Document Builder for type { type } could not be found.";
                return false;
            }
                

            try
            {
                builder.BuildFile(document);
                return true;
            }
            catch(Exception ex)
            {
                errorMsg = $"Exception: { ex.Message }.";
                return false;
            }
        }
        /// <summary>
        /// Builds a Stream containing the document model in the format in the requested document type.
        /// The Stream is to be returned via a HTTP response
        /// </summary>
        /// <param name="type"></param>
        /// <param name="document"></param>
        /// <returns>A stream for returning as a response to a HTTP request for the document</returns>
        public Stream BuildStream(DocumentBuilderType type, DocumentModel document, out string errorMsg)
        {
            errorMsg = "";
            var builder = GetBuilder(type);

            if (builder == null)
            {
                errorMsg = $"Document Builder for type { type } could not be found.";
                return null;
            }
                

            try
            {
                return builder.BuildStream(document);
            }
            catch (Exception ex)
            {
                errorMsg = $"Exception: { ex.Message }.";
                return null;
            }
        }

        protected abstract IDocumentBuilder GetBuilder(DocumentBuilderType type);

    }
}
