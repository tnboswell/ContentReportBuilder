using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using System;
using System.IO;
using Xceed.Words.NET;

namespace ContentReportBuilder.docX.Builder.Word.Elements
{
    public class WordImage : DocumentElementBuilder<DocX>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Image;

        protected override void Build(DocumentModelElement element, DocX documentSection)
        {
            ImageDataBase elementData = (ImageDataBase)element.GetData();

            Xceed.Document.NET.Image image = null;

            using (var stream = elementData.GetData())
            {
                MemoryStream mem = new MemoryStream();
                stream.CopyTo(mem);
                mem.Seek(0, SeekOrigin.Begin);
                image = documentSection.AddImage(mem);
                // Add a title
                var picture = image.CreatePicture(elementData.Height, elementData.Width);
                var imageParagraph = documentSection.InsertParagraph("");
                imageParagraph.AppendPicture(picture);
                ParagraphFormating(element.WordStyle, imageParagraph);
            }
  
        }

        
    }
}
