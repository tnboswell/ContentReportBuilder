using System;
using System.Collections.Generic;

namespace ContentReportBuilder.Models.MetaData
{
    /// <summary>
    /// Must be provided for a word or excel document
    /// </summary>
    public class DocumentMetaData : IDocumentMetaData
    {
        /// <summary>
        /// Constructor provided for a stream based builder
        /// </summary>
        /// <param name="title"></param>
        /// <param name="authors"></param>
        public DocumentMetaData(string title, List<string> authors)
        {
            Title = title;
            Author = authors;
            Created = DateTime.Now;
            Path = "";
            Keywords = new List<string>();

        }

        /// <summary>
        /// Constructor provided for a file based builder
        /// </summary>
        /// <param name="path"></param>
        /// <param name="title"></param>
        /// <param name="authors"></param>
        public DocumentMetaData(string path, string title, List<string> authors)
        {
            Path = path;
            Title = title;
            Author = authors;
            Created = DateTime.Now;
            Keywords = new List<string>();

        }
  
        public string Path { get; set; }
        public string Title { get; set; }
        public string Subject { get; set; }
        public List<string> Author { get; set; }
        public string Company { get; set; }
        public string Tags { get; set; }
        public string Comments { get; set; }
        public DateTime Created { get; set; }
        public List<string> Keywords { get; set; }
    }
}
