using DocumentFormat.OpenXml.Wordprocessing;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Builder;

namespace ContentReportBuilder.OpenXml.Builder.Word.Elements
{
    public class WordList : DocumentElementBuilderBase<Body>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.List;

        protected override void Build(DocumentModelElement element, Body documentSection)
        {

            ListData elementData = (ListData)element.GetData();

            SpacingBetweenLines sblUl = new SpacingBetweenLines() { After = "0" };  // Get rid of space between bullets  
            Indentation iUl = new Indentation() { Left = "10", Hanging = "360" };  // correct indentation  
            NumberingProperties npUl = new NumberingProperties(
                new NumberingLevelReference() { Val = 1 },
                new NumberingId() { Val = 2 }
            );

            ParagraphProperties ppUnordered = new ParagraphProperties(npUl, sblUl, iUl);
            ppUnordered.ParagraphStyleId = new ParagraphStyleId() { Val = "ListParagraph" };
            if (element.WordStyle != null)
            {


            }
            
            

            // Pargraph
            Paragraph p1 = new Paragraph();
            p1.ParagraphProperties = new ParagraphProperties(ppUnordered.OuterXml);
            p1.Append(new Run(new Text("A")));
            documentSection.Append(p1);
            Paragraph p2 = new Paragraph();
            p2.ParagraphProperties = new ParagraphProperties(ppUnordered.OuterXml);
            p2.Append(new Run(new Text("Unordored")));
            documentSection.Append(p2);
            Paragraph p3 = new Paragraph();
            p3.ParagraphProperties = new ParagraphProperties(ppUnordered.OuterXml);
            p3.Append(new Run(new Text("List")));
            documentSection.Append(p3);
        }
    }
}
