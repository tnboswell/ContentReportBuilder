using ContentReportBuilder.Builder;
using ContentReportBuilder.Models.Styles;
using iTextSharp.text;
using OxyPlot.Core.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ContentReportBuilder.ITextSharp.Builder
{
    public abstract class DocumentElementBuilder<T> : DocumentElementBuilderBase<T>, IDocumentElementBuilder
    {
        protected void ParagraphFormating(IWordElementStyle style, Paragraph paragraph)
        {
            var font = style.Font;
            paragraph.Font.Size = font.Size;
            paragraph.Font.Color = new BaseColor(font.Colour);
            paragraph.Font.SetFamily(font.Font);
            var fontStyle = new StringBuilder();
            if (font.IsBold)
                fontStyle.Append("bold, ");

            if (font.IsItalics)
                fontStyle.Append("italic, ");

            if (font.IsItalics)
                fontStyle.Append("underline, ");

            if (font.IsStrikeThrough)
                fontStyle.Append("strike, ");

            if (font.IsUnderlined)
                fontStyle.Append("underline, ");

            if (fontStyle.Length == 0)
                fontStyle.Append("normal, ");

            fontStyle.Length--;

            paragraph.Font.SetStyle(fontStyle.ToString());
            paragraph.Alignment = GetHorizontalAlignment(style.HorizontalAlignment);
            paragraph.FirstLineIndent = (float)style.Indentation.FirstLine;
            paragraph.IndentationLeft = (float)style.Indentation.Hanging;
            paragraph.SpacingAfter = (float)style.Spacing.Top;
            paragraph.SpacingBefore = (float)style.Spacing.Bottom;
        }

        protected int GetHorizontalAlignment(HorizontalAlignment alignment)
        {
            switch (alignment)
            {
                case HorizontalAlignment.Centre:
                    return 1;
                case HorizontalAlignment.Left:
                    return 0;
                case HorizontalAlignment.Right:
                    return 2;
                default:
                    return 0;
            }
        }

        protected void WritePlotToDocument(Document documentSection, OxyPlot.PlotModel plotModel, int height, int width)
        {
            using (var stream = new MemoryStream())
            {

                var pngExporter = new PngExporter { Width = 240, Height = 160, Background = OxyPlot.OxyColors.White };
                pngExporter.Export(plotModel, stream);
                // try microcharts or livecharts if you have no luck with oxyplot
                stream.Seek(0, SeekOrigin.Begin);

                var image = Image.GetInstance(stream);
                documentSection.Add(image);

            }
        }

    }
}
