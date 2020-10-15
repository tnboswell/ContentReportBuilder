
namespace ContentReportBuilder.Models.Styles
{
    /// <summary>
    /// Style used to create excel based elements
    /// </summary>
    public class ExcelElementStyle : IExcelElementStyle
    {
        /// <summary>
        /// Default Constructor
        /// Sets the style elements to default settings which can be overriden
        /// </summary>
        public ExcelElementStyle()
        {
            HorizontalAlignment = HorizontalAlignment.Centre;
            VeritcalAlignment = VeritcalAlignment.Middle;
            Font = new FontStyle();
            Cell = new CellStyle();
        }

        public CellStyle Cell { get; set; }
        public HorizontalAlignment HorizontalAlignment { get; set; }
        public VeritcalAlignment VeritcalAlignment { get; set; }
        public FontStyle Font { get; set; }
    }
}
