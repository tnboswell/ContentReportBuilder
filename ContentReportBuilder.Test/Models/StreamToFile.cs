using ContentReportBuilder.Models;
using ContentReportBuilder.Models.MetaData;
using System.IO;
using System.Linq;


namespace ContentReportBuilder.Test.Models
{
    public class StreamToFile
    {
        public static void WriteStream(DocumentBuilderType docType, DocumentModel docModel, Stream stream)
        {
            using (MemoryStream memoryStream = (MemoryStream)stream)
            {
                byte[] content = memoryStream.ToArray();
                using (FileStream fs = File.Create(Filename(docModel.MetaData, docType)))
                {
                    fs.Write(content, 0, (int)content.Length);
                }
            }
        }
        protected static string Filename(IDocumentMetaData metaData, DocumentBuilderType docType)
        {
            if (string.IsNullOrWhiteSpace(metaData.Path))
            {
                return $"{metaData.Title.Split('.').First()}.{FileExtension(docType)}";
            }

            if (!Directory.Exists(metaData.Path))
            {
                Directory.CreateDirectory(metaData.Path);
            }

            return $"{metaData.Path}/{metaData.Title.Split('.').First()}.{FileExtension(docType)}";
        }

        protected static string FileExtension(DocumentBuilderType type)
        {
            switch (type)
            {
                case DocumentBuilderType.Word:
                    return "docx";
                case DocumentBuilderType.Excel:
                    return "xlsx";
                case DocumentBuilderType.Pdf:
                    return "pdf";
                default:
                    return "";
            }
        }
    }
}
