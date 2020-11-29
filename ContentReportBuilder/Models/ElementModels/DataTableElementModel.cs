using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.Styles;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Drawing;

namespace ContentReportBuilder.Models.ElementModels
{
    /// <summary>
    /// Implemented Table Element Model
    /// </summary>
    public class DataTableElementModel : DocumentModelElement
    {
        /// <summary>
        /// Default Constructor for Table element model
        /// </summary>
        /// <param name="data"></param>
        /// <param name="wordStyle"></param>
        /// <param name="excelStyle"></param>
        public DataTableElementModel(TableData data, Dictionary<int, IExcelElementStyle> colStyle = null, IWordElementStyle defaultWordStyle = null, IExcelElementStyle defaultExcelStyle = null) :
            base(DocumentElementType.Table, JObject.FromObject(data), defaultWordStyle, defaultExcelStyle)
        {
            ColStyle = colStyle;
        }
        public Color HeadingColour { get; set; }
        public Dictionary<int, IExcelElementStyle> ColStyle { get; set; }
        /// <summary>
        /// Type safe validate and get the chart data for the element builder
        /// </summary>
        /// <returns></returns>
        public override object GetData()
        {
            return Data.ToObject<TableData>();
        }
    }
}
