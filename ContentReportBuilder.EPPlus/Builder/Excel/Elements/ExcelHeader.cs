using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using OfficeOpenXml;

namespace ContentReportBuilder.EPPlus.Builder.Excel.Elements
{
    public class ExcelHeader : DocumentElementBuilder<ExcelWorksheet>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Header;

        protected override void Build(DocumentModelElement element, ExcelWorksheet documentSection)
        {
            HeaderData elementData = (HeaderData)element.GetData();

            var header = documentSection.HeaderFooter.FirstHeader;
            var evenHeader = documentSection.HeaderFooter.EvenHeader;
            var oddHeader = documentSection.HeaderFooter.OddHeader;

            if (elementData.Image != null)
            {

                using (System.Drawing.Image image = System.Drawing.Image.FromStream(elementData.Image.GetData()))
                {
                    PictureAlignment pictureAlignment = PictureAlignment.Centered;
                    if (element.ExcelStyle.HorizontalAlignment == Models.Styles.HorizontalAlignment.Left)
                        pictureAlignment = PictureAlignment.Left;
                    else if (element.ExcelStyle.HorizontalAlignment == Models.Styles.HorizontalAlignment.Right)
                        pictureAlignment = PictureAlignment.Right;

                    header.InsertPicture(image, pictureAlignment);


                }
            }
            else if (elementData.Paragraph != null)
            {
                if (element.ExcelStyle.HorizontalAlignment == Models.Styles.HorizontalAlignment.Left)
                    header.LeftAlignedText = elementData.Paragraph.Text;
                else if (element.ExcelStyle.HorizontalAlignment == Models.Styles.HorizontalAlignment.Right)
                    header.RightAlignedText = elementData.Paragraph.Text;
                else
                    header.CenteredText = elementData.Paragraph.Text;
            }
            else
                return;

            evenHeader = header;
            oddHeader = header;
        }
    }
}
