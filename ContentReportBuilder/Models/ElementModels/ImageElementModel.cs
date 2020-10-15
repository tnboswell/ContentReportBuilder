using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.Styles;
using Newtonsoft.Json.Linq;


namespace ContentReportBuilder.Models.ElementModels
{
    /// <summary>
    /// Implemented image Element Model
    /// </summary>
    public class ImageElementModel : DocumentModelElement
    {
        /// <summary>
        /// Default Constructor for Chart element model
        /// </summary>
        /// <param name="data"></param>
        /// <param name="wordStyle"></param>
        /// <param name="excelStyle"></param>
        public ImageElementModel(ImageDataBase data, IWordElementStyle wordStyle = null, IExcelElementStyle excelStyle = null) :
            base(DocumentElementType.Image, JObject.FromObject(data), wordStyle, excelStyle)
        {

        }

        /// <summary>
        /// Type safe validate and get the chart data for the element builder
        /// </summary>
        /// <returns></returns>
        public override object GetData()
        {

            if (Data.ContainsKey("URL"))
            {
                return Data.ToObject<WebImageData>();
            }
            else if (Data.ContainsKey("FilePath"))
            {
                return Data.ToObject<LocalImageData>();
            }

            throw new System.Exception();
        }
    }
}
