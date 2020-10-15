
namespace ContentReportBuilder.Models.Styles
{
    /// <summary>
    /// Style used for a word document
    /// </summary>
    public class DocumentStyle : IDocumentStyle
    {
        public DocumentStyle(bool excel = false)
        {
            if (excel)
            {
                Margins = new MarginStyle(1,1,1,1);
            }
            else
            {
                Margins = new MarginStyle();
            }
            
            HeaderHeight = 0;
            FooterHeight = 0;
        }
        public MarginStyle Margins { get; set; }
        public double HeaderHeight { get; set; }
        public double FooterHeight { get; set; }
    }
}
