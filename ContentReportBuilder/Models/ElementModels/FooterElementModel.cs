using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.Styles;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Text;

namespace ContentReportBuilder.Models.ElementModels
{
    public class FooterElementModel : DocumentModelElement
    {
        public FooterElementModel(FooterData data, IWordElementStyle wordStyle = null, IExcelElementStyle excelStyle = null) 
            : base(DocumentElementType.Footer, JObject.FromObject(data), wordStyle, excelStyle)
        {
        }

        public override object GetData()
        {
            var output = new FooterData();
            var paragraphToken = Data["Paragraph"];
            if (paragraphToken.Type != JTokenType.Null)
            {
                output.Paragraph = paragraphToken.ToObject<ParagraphData>();
            }

            var imageToken = Data["Image"];
            if (imageToken.Type == JTokenType.Null)
                return output;

            if (((JObject)imageToken).ContainsKey("URL"))
            {
                output.Image = imageToken.ToObject<WebImageData>();
            }
            else if (((JObject)imageToken).ContainsKey("FilePath"))
            {
                output.Image = imageToken.ToObject<LocalImageData>();
            }

            return output; //Data.ToObject<FooterData>();
        }
    }
}
