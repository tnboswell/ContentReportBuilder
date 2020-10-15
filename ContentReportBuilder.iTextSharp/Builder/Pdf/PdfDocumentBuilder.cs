using ContentReportBuilder.Builder;
using ContentReportBuilder.ITextSharp.Builder.Pdf.Sections;
using ContentReportBuilder.Models;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ContentReportBuilder.ITextSharp.Builder.Pdf
{
    public class PdfDocumentBuilder : DocumentBuilderBase<Document, Document, Document>, IDocumentBuilder
    {
        public PdfDocumentBuilder() : base(new PdfSectionBuilder())
        {
        }
        public override DocumentBuilderType Type => DocumentBuilderType.Pdf;

        public override Document Build(DocumentModel document, Document package)
        {
            foreach (var section in document.DocumentSections ?? new List<DocumentModelSection>())
            {
                package.SetMargins(
                    (float)document.Style.Margins.Left,
                    (float)document.Style.Margins.Right,
                    (float)document.Style.Margins.Top,
                    (float)document.Style.Margins.Bottom);
                _sectionBuilder.Build(Type, section, package);
            }

            return package;
        }

        public override void BuildFile(DocumentModel document)
        {

            using (MemoryStream memStream = new MemoryStream())
            {
                
                var marginLeft = document.Style.Margins.Left;
                var marginRight = document.Style.Margins.Right;
                var marginTop = document.Style.Margins.Top;
                var marginBottom = document.Style.Margins.Bottom;
                Document package = new Document(PageSize.A4, (float)marginLeft, (float)marginRight, (float)marginTop, (float)marginBottom);
                PdfWriter writer = PdfWriter.GetInstance(package, memStream);

                var authors = new StringBuilder();
                foreach (var author in document.MetaData.Author)
                {
                    authors.Append(author);
                    authors.Append(", ");
                }
                authors.Length--;
                authors.Length--;

                package.AddAuthor(authors.ToString());
                package.AddKeywords(document.MetaData.Tags);
                package.AddCreator(Constants.ApplicationName);
                package.AddSubject(document.MetaData.Subject);
                package.AddTitle(document.MetaData.Title);
                package.AddCreationDate();
             

                package.Open();
                Build(document, package);
                package.Close();
                writer.Close();

                byte[] content = memStream.ToArray();

                // Write out PDF from memory stream.
                using (FileStream fs = File.Create(Filename(document.MetaData)))
                {
                    fs.Write(content, 0, (int)content.Length);
                }

            }
        }

        public override Stream BuildStream(DocumentModel document)
        {
            Stream mem = new MemoryStream();
            var marginLeft = document.Style.Margins.Left;
            var marginRight = document.Style.Margins.Right;
            var marginTop = document.Style.Margins.Top;
            var marginBottom = document.Style.Margins.Bottom;
            // Create an instance of the document class which represents the PDF document itself.  
            using (Document package = new Document(PageSize.A4, (float)marginLeft, (float)marginRight, (float)marginTop, (float)marginBottom))
            {
                // Create an instance to the PDF file by creating an instance of the PDF   
                // Writer class using the document and the filestrem in the constructor.  

                PdfWriter writer = PdfWriter.GetInstance(package, mem);

                var authors = new StringBuilder();
                foreach (var author in document.MetaData.Author)
                {
                    authors.Append(author);
                    authors.Append(", ");
                }
                authors.Length--;
                authors.Length--;

                package.AddAuthor(authors.ToString());
                package.AddKeywords(document.MetaData.Tags);
                package.AddCreator(Constants.ApplicationName);
                package.AddSubject(document.MetaData.Subject);
                package.AddTitle(document.MetaData.Title);
                package.AddCreationDate();


                // Open the document to enable you to write to the document  
                package.Open();
                Build(document, package);
                package.Close();
                writer.Close();
                return mem;
            }

        }
    }
}
