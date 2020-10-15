using DocumentFormat.OpenXml.Wordprocessing;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Builder;

namespace ContentReportBuilder.OpenXml.Builder.Word.Elements
{
    public class WordParagraph : DocumentElementBuilderBase<Body>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.Paragraph;

       

        protected override void Build(DocumentModelElement element, Body documentSection)
        {

            RunProperties runProperties = new RunProperties();
            if (element.WordStyle != null)
            {
                if (element.WordStyle.Font.IsBold)
                    runProperties.Bold = new Bold();

                if (element.WordStyle.Font.IsItalics)
                    runProperties.Italic = new Italic();

                if (element.WordStyle.Font.IsStrikeThrough)
                    runProperties.Strike = new Strike();

                if (element.WordStyle.Font.IsUnderlined)
                    runProperties.Underline = new Underline();

                runProperties.FontSize = new FontSize { Val = element.WordStyle.Font.Size.ToString() };
                runProperties.Color = new Color { Val = element.WordStyle.Font.HexConverter() };

            }


            Paragraph paragraph = new Paragraph();
            
            Run run = new Run();
            run.Append(runProperties);

            ParagraphData elementData = (ParagraphData)element.GetData();

            Text text = new Text(elementData.Text);
            run.Append(text);
            paragraph.Append(run);
            documentSection.Append(paragraph);
        }
    }
}
