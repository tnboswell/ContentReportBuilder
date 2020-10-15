using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ContentReportBuilder.Models.ElementDataModels
{
    public class TableDataAdapter
    {
        public List<string> Headers(object row)
        {
            List<string> headers = new List<string>();
            var objectType = row.GetType();
            if (objectType.Name == "JObject")
            {
                foreach (JProperty field in (row as JToken))
                {
                    headers.Add(field.Name);
                }
            }
            return headers;
        }

        public string ReadableHeader(string header)
        {
            return string.Concat(header.Select(x => Char.IsUpper(x) ? " " + x : x.ToString())).TrimStart(' ');
        }

        public string CellValue(object row, string key)
        {
            var objectType = row.GetType();
            if (objectType.Name == "JObject" && (row as JObject).ContainsKey(key))
            {
                // get value of property: public double Number
                return (row as JObject)[key].ToString();
            }
            throw new Exception();

        }
    }
}
