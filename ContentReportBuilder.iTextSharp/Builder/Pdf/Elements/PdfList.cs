using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using iTextSharp.text;
using System;

namespace ContentReportBuilder.ITextSharp.Builder.Pdf.Elements
{
    //https://www.mikesdotnetting.com/article/83/lists-with-itextsharp
    public class PdfList : DocumentElementBuilderBase<Document>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.List;

        protected override void Build(DocumentModelElement element, Document documentSection)
        {
            ListData elementData = (ListData)element.GetData();
            
            var list = new List(elementData.Ordered);
            list.SymbolIndent = 20f; // space between symbol and text
            list.Autoindent = true;
            //list.IndentationLeft = 20;// (float)element.WordStyle.Indentation.Hanging;

            list.Autoindent = true;

            if (!elementData.Ordered)
            {
                list.SetListSymbol("\u2022");
            }
            
            foreach(var item in elementData.Items ?? new System.Collections.Generic.List<ListItemData>())
            {
                var listItem = new ListItem(item.Text);
                list.Add(listItem);
            }
            documentSection.Add(list);

        }
    }
}
