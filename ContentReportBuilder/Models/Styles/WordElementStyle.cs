
namespace ContentReportBuilder.Models.Styles
{
    /// <summary>
    /// Styles used to generated a word based element
    /// </summary>
    public class WordElementStyle : IWordElementStyle
    {
        /// <summary>
        /// Default Constructor
        /// Sets the style elements to default settings which can be overriden
        /// </summary>
        public WordElementStyle()
        {
            Indentation = new IndentationStyle();
            Spacing = new SpacingStyle();
            Font = new FontStyle();
            HorizontalAlignment = HorizontalAlignment.Left;
            VeritcalAlignment = VeritcalAlignment.Top;
        }
        public IndentationStyle Indentation { get; set; }
        public SpacingStyle Spacing { get; set; }
        public HorizontalAlignment HorizontalAlignment { get; set; }
        public VeritcalAlignment VeritcalAlignment { get; set; }
        public FontStyle Font { get; set; }
    }
}
