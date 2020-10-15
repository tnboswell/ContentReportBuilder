using ContentReportBuilder.Builder;
using ContentReportBuilder.Models.Styles;
using System;
using System.Collections.Generic;
using System.Text;
// https://github.com/xceedsoftware/DocX/tree/master/Xceed.Words.NET.Examples/Samples
namespace ContentReportBuilder.docX.Builder
{
    public abstract class DocumentElementBuilder<T> : DocumentElementBuilderBase<T>, IDocumentElementBuilder
    {
        protected void ParagraphFormating(IWordElementStyle style, Xceed.Document.NET.Paragraph paragraph)
        {
            var font = style.Font;
            paragraph.Font(font.Font);
            paragraph.FontSize(font.Size);

            paragraph.Bold(font.IsBold);
            paragraph.Italic(font.IsItalics);
            if (font.IsStrikeThrough)
            {
                paragraph.StrikeThrough(Xceed.Document.NET.StrikeThrough.strike);
            }
            if (font.IsUnderlined)
            {
                paragraph.UnderlineStyle(Xceed.Document.NET.UnderlineStyle.singleLine);
            }
            paragraph.Color(font.Colour);

            paragraph.Alignment = GetHorizontalAlignment(style.HorizontalAlignment);

            var indentation = style.Indentation;

            paragraph.IndentationFirstLine = (float)indentation.FirstLine;
            paragraph.IndentationHanging = (float)indentation.Hanging;

            var spacing = style.Spacing;

            paragraph.SpacingBefore(spacing.Top);
            paragraph.SpacingAfter(spacing.Bottom);
            paragraph.SpacingLine(spacing.LineSpacing);
        }

        protected Xceed.Document.NET.Alignment GetHorizontalAlignment(HorizontalAlignment alignment)
        {
            switch (alignment)
            {
                case HorizontalAlignment.Centre:
                    return Xceed.Document.NET.Alignment.center;
                case HorizontalAlignment.Left:
                    return Xceed.Document.NET.Alignment.left;
                case HorizontalAlignment.Right:
                    return Xceed.Document.NET.Alignment.right;
                case HorizontalAlignment.Justified:
                    return Xceed.Document.NET.Alignment.both;
                default:
                    return Xceed.Document.NET.Alignment.center;
            }

        }
    }
}
