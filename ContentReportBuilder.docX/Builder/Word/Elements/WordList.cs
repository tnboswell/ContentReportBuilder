using ContentReportBuilder.Builder;
using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using System.Collections.Generic;
using Xceed.Words.NET;

namespace ContentReportBuilder.docX.Builder.Word.Elements
{
    public class WordList : DocumentElementBuilderBase<DocX>, IDocumentElementBuilder
    {
        public override DocumentElementType Type => DocumentElementType.List;

        protected override void Build(DocumentModelElement element, DocX documentSection)
        {
            ListData elementData = (ListData)element.GetData();
            Xceed.Document.NET.List list = null;
            bool first = true;
            foreach(var listItem in elementData.Items ?? new List<ListItemData>())
            {

                if (first && elementData.Ordered)
                {
                    list = documentSection.AddList(listItem.Text, 0, Xceed.Document.NET.ListItemType.Numbered, 1);
                    first = false;
                }
                else if(first)
                {
                    list = documentSection.AddList(listItem.Text, 0, Xceed.Document.NET.ListItemType.Bulleted);
                    first = false;
                }
                else
                {
                    documentSection.AddListItem(list, listItem.Text);
                }
            }
            var font = element.WordStyle.Font;
            documentSection.InsertList(list, new Xceed.Document.NET.Font(font.Font), font.Size);
        }
    }
}
