using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.OpenXml.Builder.Word.Sections;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.IO;

namespace ContentReportBuilder.OpenXml.Builder.Word
{
    public class WordDocumentBuilder : DocumentBuilderBase<Body, WordprocessingDocument, MainDocumentPart>, IDocumentBuilder
    {
        public WordDocumentBuilder(): base(new WordSectionBuilder())
        {

        }
        
        public override DocumentBuilderType Type => DocumentBuilderType.Word;

        public override Stream BuildStream(DocumentModel document)
        {
            Stream mem = new MemoryStream();
            using (WordprocessingDocument package = WordprocessingDocument.Create(mem, WordprocessingDocumentType.Document, true))
            {
                Build(document, package);
                return mem;
            }
        }

        public override void BuildFile(DocumentModel document)
        {
            using (WordprocessingDocument package = WordprocessingDocument.Create(Filename(document.MetaData), WordprocessingDocumentType.Document))
            {
                MainDocumentPart workbookPart = Build(document, package);
                workbookPart.Document.Save();
            }
        }

        public override MainDocumentPart Build(DocumentModel document, WordprocessingDocument package)
        {
            // Add a main document part. 
            MainDocumentPart mainPart = package.AddMainDocumentPart();
            // Create the document structure and add some text.
            mainPart.Document = new Document();
            Body docBody = new Body();

            foreach (var section in document.DocumentSections ?? new List<DocumentModelSection>())
            {
                _sectionBuilder.Build(Type, section, docBody);
            }

            return mainPart;
        }
    }
}
