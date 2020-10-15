using Newtonsoft.Json.Linq;
using ContentReportBuilder.Models.Styles;

namespace ContentReportBuilder.Models
{
    public abstract class DocumentModelElement
    {
        /// <summary>
        /// Constructor for the base class for document element models. Element models must be provided an Element Type and Json data containing the elements value/data.
        /// Styles are optionally provided to enable the word and/or excel versions of the element to be styled to produce a final output.
        /// </summary>
        /// <param name="type"><c>DocumentElementType</c> defines the supported document element types</param>
        /// <param name="data">Element data as a Json Object</param>
        /// <param name="wordStyle">Element style for word documents</param>
        /// <param name="excelStyle">Elmenet style for excel documents</param>
        public DocumentModelElement(DocumentElementType type, JObject data, IWordElementStyle wordStyle = null, IExcelElementStyle excelStyle = null)
        {
            Type = type;
            Data = data;
            WordStyle = wordStyle ?? new WordElementStyle();
            ExcelStyle = excelStyle ?? new ExcelElementStyle();
        }
        public DocumentElementType Type { get; protected set; }
        public IWordElementStyle WordStyle { get; set; }
        public IExcelElementStyle ExcelStyle { get; set; }
        protected JObject Data { get; set; }

        public abstract object GetData();

    }
}
