using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.Styles;
using Newtonsoft.Json.Linq;

namespace ContentReportBuilder.Models.ElementModels
{
    /// <summary>
    /// Implemented List Element Model
    /// </summary>
    public class ListElementModel : DocumentModelElement
    {
        /// <summary>
        /// Default Constructor for Chart element model
        /// </summary>
        /// <param name="data"></param>
        /// <param name="wordStyle"></param>
        /// <param name="excelStyle"></param>
        public ListElementModel(ListData data, IWordElementStyle wordStyle = null, IExcelElementStyle excelStyle = null) :
            base(DocumentElementType.List, JObject.FromObject(data), wordStyle, excelStyle)
        {

        }

        /// <summary>
        /// Type safe validate and get the chart data for the element builder
        /// </summary>
        /// <returns></returns>
        public override object GetData()
        {
            return Data.ToObject<ListData>();
        }
    }
}
