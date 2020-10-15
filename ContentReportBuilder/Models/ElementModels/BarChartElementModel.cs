using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.Styles;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Text;

namespace ContentReportBuilder.Models.ElementModels
{
    public class BarChartElementModel : DocumentModelElement
    {
        public BarChartElementModel(ChartData data, IWordElementStyle wordStyle = null, IExcelElementStyle excelStyle = null) :
            base(DocumentElementType.BarChart, JObject.FromObject(data), wordStyle, excelStyle)
        {

        }

        public override object GetData()
        {
            return Data.ToObject<ChartData>();
        }
    }
}
