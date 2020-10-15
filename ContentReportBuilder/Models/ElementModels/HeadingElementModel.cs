using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.Styles;
using Newtonsoft.Json.Linq;

namespace ContentReportBuilder.Models.ElementModels
{
    /// <summary>
    /// Implemented Heading Element Model
    /// </summary>
    public class HeadingElementModel : DocumentModelElement
    {
        /// <summary>
        /// Default Constructor for Chart element model
        /// </summary>
        /// <param name="data"></param>
        /// <param name="wordStyle"></param>
        /// <param name="excelStyle"></param>
        public HeadingElementModel(HeadingData data, IWordElementStyle wordStyle = null, IExcelElementStyle excelStyle = null) :
            base(DocumentElementType.Heading, JObject.FromObject(data), wordStyle, excelStyle)
        {

        }

        /// <summary>
        /// Type safe validate and get the chart data for the element builder
        /// </summary>
        /// <returns></returns>
        public override object GetData()
        {
            return Data.ToObject<HeadingData>();
        }
    }
}
