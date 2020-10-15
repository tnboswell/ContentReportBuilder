using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using iTextSharp.text;
using System.IO;

namespace ContentReportBuilder.ITextSharp.Builder.Pdf.Elements
{
    public class PdfImage : DocumentElementBuilderBase<Document>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Image;

        protected override void Build(DocumentModelElement element, Document documentSection)
        {
            ImageDataBase elementData = (ImageDataBase)element.GetData();
            var image = Image.GetInstance(elementData.GetData());
            documentSection.Add(image);
        }
    }
}
