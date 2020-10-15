using ContentReportBuilder.Models;
using ContentReportBuilder.Models.MetaData;
using System.IO;
using System.Linq;


namespace ContentReportBuilder.Builder
{
    public abstract class DocumentBuilderBase<SectionBuilderType, DocumentPackageType, BuildReturnType>: IDocumentBuilder
    {
        protected readonly IDocumentSectionBuilder<SectionBuilderType> _sectionBuilder;
        public abstract DocumentBuilderType Type { get; }
        public DocumentBuilderBase(IDocumentSectionBuilder<SectionBuilderType> sectionBuilder)
        {
            _sectionBuilder = sectionBuilder;
        }

        public abstract BuildReturnType Build(DocumentModel document, DocumentPackageType package);

        /// <summary>
        /// Returns a filename with the builders file extension. 
        /// If a fullstop is used in the meta data filename the fullstop and all characters after a stripped from the name.
        /// </summary>
        /// <param name="metaData"></param>
        /// <returns></returns>
        public string Filename(IDocumentMetaData metaData)
        {
            if (string.IsNullOrWhiteSpace(metaData.Path))
            {
                return $"{metaData.Title.Split('.').First()}.{FileExtension()}";
            }

            if (!Directory.Exists(metaData.Path))
            {
                Directory.CreateDirectory(metaData.Path);
            }

            return $"{metaData.Path}/{metaData.Title.Split('.').First()}.{FileExtension()}";
        }
        
        public abstract Stream BuildStream(DocumentModel document);
        public abstract void BuildFile(DocumentModel document);

        protected string FileExtension()
        {
            switch (Type)
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
