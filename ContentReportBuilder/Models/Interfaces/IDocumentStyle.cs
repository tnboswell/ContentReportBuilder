
namespace ContentReportBuilder.Models.Styles
{
    /// <summary>
    /// Provides document styles used for word documents
    /// Excel does not have any document level sytles and will be ignored if provided
    /// </summary>
    public interface IDocumentStyle
    {
        MarginStyle Margins { get; set; }
        double HeaderHeight { get; set; }
        double FooterHeight { get; set; }
    }
}
