

using ContentReportBuilder.Builder;
using ContentReportBuilder.docX.Builder.Word.Sections;
using ContentReportBuilder.Models;
using System.Collections.Generic;
using System.IO;
using Xceed.Words.NET;

namespace ContentReportBuilder.docX.Builder.Word
{
    public class WordDocumentBuilder : DocumentBuilderBase<DocX, DocX, DocX>, IDocumentBuilder
    {
        public WordDocumentBuilder() : base(new WordSectionBuilder())
        {
        }
        public override DocumentBuilderType Type => DocumentBuilderType.Word;

        public override DocX Build(DocumentModel document, DocX package)
        {
               
            foreach (var section in document.DocumentSections ?? new List<DocumentModelSection>())
            {
                package.MarginTop = (float)document.Style.Margins.Top;
                package.MarginBottom = (float)document.Style.Margins.Bottom;
                package.MarginLeft = (float)document.Style.Margins.Left;
                package.MarginRight = (float)document.Style.Margins.Right;
                package.MarginHeader = (float)document.Style.HeaderHeight;
                package.MarginFooter = (float)document.Style.FooterHeight;
                _sectionBuilder.Build(Type, section, package);
            }

            return package;
        }

        public override void BuildFile(DocumentModel document)
        {
            using (DocX package = DocX.Create(Filename(document.MetaData)))
            {
                Build(document, package);
                package.Save();
            }
                
        }

        public override Stream BuildStream(DocumentModel document)
        {
            Stream mem = new MemoryStream();
            using (DocX package = DocX.Create(mem))
            {
                Build(document, package);
                package.Save();
                return mem;
            }
        }
    }
}
