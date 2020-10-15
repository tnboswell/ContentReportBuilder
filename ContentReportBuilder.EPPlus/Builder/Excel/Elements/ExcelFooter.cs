using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;

namespace ContentReportBuilder.EPPlus.Builder.Excel.Elements
{
    public class ExcelFooter : DocumentElementBuilder<ExcelWorksheet>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Footer;

        protected override void Build(DocumentModelElement element, ExcelWorksheet documentSection)
        {
            FooterData elementData = (FooterData)element.GetData();
            documentSection.HeaderFooter.differentFirst = true;
            documentSection.HeaderFooter.differentOddEven = true;
            var footer = documentSection.HeaderFooter.FirstFooter;
            var evenFooter = documentSection.HeaderFooter.EvenFooter;
            var oddFooter = documentSection.HeaderFooter.OddFooter;

            if (elementData.Image != null)
            {

                using (System.Drawing.Image image = System.Drawing.Image.FromStream(elementData.Image.GetData()))
                {
                    PictureAlignment pictureAlignment = PictureAlignment.Centered;
                    if (element.ExcelStyle.HorizontalAlignment == Models.Styles.HorizontalAlignment.Left)
                        pictureAlignment = PictureAlignment.Left;
                    else if (element.ExcelStyle.HorizontalAlignment == Models.Styles.HorizontalAlignment.Right)
                        pictureAlignment = PictureAlignment.Right;

                    footer.InsertPicture(image, pictureAlignment);
                    evenFooter.InsertPicture(image, pictureAlignment);
                    oddFooter.InsertPicture(image, pictureAlignment);

                }
            }
            else if (elementData.Paragraph != null)
            {
                if (element.ExcelStyle.HorizontalAlignment == Models.Styles.HorizontalAlignment.Left)
                    footer.LeftAlignedText = elementData.Paragraph.Text;
                else if (element.ExcelStyle.HorizontalAlignment == Models.Styles.HorizontalAlignment.Right)
                    footer.RightAlignedText = elementData.Paragraph.Text;
                else
                    footer.CenteredText = elementData.Paragraph.Text;
            }
            else
                return;

            evenFooter = footer;
            oddFooter = footer;
        }

    }
}
