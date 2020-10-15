using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.Styles;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Text;

namespace ContentReportBuilder.Models.ElementModels
{
    public class PieChartElementModel : DocumentModelElement
    {
        public PieChartElementModel(ChartData data, IWordElementStyle wordStyle = null, IExcelElementStyle excelStyle = null) :
            base(DocumentElementType.PieChart, JObject.FromObject(data), wordStyle, excelStyle)
        {

        }
        public override object GetData()
        {
            return Data.ToObject<ChartData>();
        }
    }
}
