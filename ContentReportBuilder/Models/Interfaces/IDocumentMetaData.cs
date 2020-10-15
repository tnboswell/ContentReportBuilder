using System;
using System.Collections.Generic;

namespace ContentReportBuilder.Models.MetaData
{
    /// <summary>
    /// Must be provided as meta data is used to create both word and excel documents
    /// </summary>
    public interface IDocumentMetaData
    {
        /// <value> Used by the builder to write a built document to the file system </value>
        string Path { get; set; }
        /// <value> Title of the document without the file extension </value>
        string Title { get; set; }
        string Subject { get; set; }
        List<string> Author { get; set; }

        List<string> Keywords { get; set; }
        string Company { get; set; }
        string Tags { get; set; }
        string Comments { get; set; }
        DateTime Created { get; set; }


    }
}
