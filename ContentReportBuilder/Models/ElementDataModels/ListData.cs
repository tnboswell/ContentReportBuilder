using System.Collections.Generic;

namespace ContentReportBuilder.Models.ElementDataModels
{
    /// <summary>
    /// List element model data
    /// </summary>
    public class ListData
    {
        public ListData()
        {

        }
        public ListData(bool ordered = true)
        {
            Ordered = ordered;
        }
        public bool Ordered {get;set;}
        public List<ListItemData> Items { get; set;}
    }
}
