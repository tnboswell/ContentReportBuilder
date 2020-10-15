using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using System.IO;
using Xceed.Words.NET;

namespace ContentReportBuilder.docX.Builder.Word.Elements
{
    public class WordFooter : DocumentElementBuilder<DocX>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Footer;

        protected override void Build(DocumentModelElement element, DocX documentSection)
        {
            FooterData elementData = (FooterData)element.GetData();
            documentSection.AddFooters();
            documentSection.DifferentOddAndEvenPages = true;
            if (elementData.Image != null)
            {
                using (var stream = elementData.Image.GetData())
                {
                    MemoryStream mem = new MemoryStream();
                    stream.CopyTo(mem);
                    mem.Seek(0, SeekOrigin.Begin);
                    Xceed.Document.NET.Image image = documentSection.AddImage(mem);
                    var picture = image.CreatePicture();
                    var oddParagraph = documentSection.Footers.Odd.InsertParagraph("");
                    oddParagraph.AppendPicture(picture);
                    var evenParagraph = documentSection.Footers.Even.InsertParagraph("");
                    evenParagraph.AppendPicture(picture);
                }
            }
            else if (elementData.Paragraph != null)
            {
                var paragragh = documentSection.Footers.Odd.InsertParagraph(elementData.Paragraph.Text);
                documentSection.Footers.Even.InsertParagraph(elementData.Paragraph.Text);
                ParagraphFormating(element.WordStyle, paragragh);
            }
            else
                return;

        }
    }
}
