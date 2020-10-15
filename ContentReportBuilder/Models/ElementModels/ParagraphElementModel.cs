using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.Styles;
using Newtonsoft.Json.Linq;

namespace ContentReportBuilder.Models.ElementModels
{
    /// <summary>
    /// Implemented Paragraph Element Model
    /// </summary>
    public class ParagraphElementModel: DocumentModelElement
    {
        /// <summary>
        /// Default Constructor for paragraph element model
        /// </summary>
        /// <param name="data"></param>
        /// <param name="wordStyle"></param>
        /// <param name="excelStyle"></param>
        public ParagraphElementModel(ParagraphData data, IWordElementStyle wordStyle = null, IExcelElementStyle excelStyle = null): 
            base(DocumentElementType.Paragraph, JObject.FromObject(data), wordStyle, excelStyle)
        {

        }

        /// <summary>
        /// Type safe validate and get the paragraph data for the element builder
        /// </summary>
        /// <returns></returns>
        public override object GetData()
        {
            return Data.ToObject<ParagraphData>();
        }
    }
}
