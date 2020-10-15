using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using System.IO;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace ContentReportBuilder.docX.Builder.Word.Elements
{
    public class WordHeader : DocumentElementBuilder<DocX>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Header;
        protected override void Build(DocumentModelElement element, DocX documentSection)
        {
            HeaderData elementData = (HeaderData)element.GetData();
            documentSection.AddHeaders();
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
                    var oddParagraph = documentSection.Headers.Odd.InsertParagraph("");
                    oddParagraph.AppendPicture(picture);
                    var evenParagraph = documentSection.Headers.Even.InsertParagraph("");
                    evenParagraph.AppendPicture(picture);
                }
            }
            else if (elementData.Paragraph != null)
            {
                var paragragh = documentSection.Headers.Odd.InsertParagraph(elementData.Paragraph.Text);
                documentSection.Headers.Even.InsertParagraph(paragragh);
                ParagraphFormating(element.WordStyle, paragragh);
            }
            else
                return;
        }
    }
}
