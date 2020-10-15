
namespace ContentReportBuilder.Models.Styles
{
    /// <summary>
    /// Base element styles used by both excel and word elements
    /// </summary>
    public interface IDocumentElementStyle
    {
        HorizontalAlignment HorizontalAlignment { get; set; }
        VeritcalAlignment VeritcalAlignment { get; set; }
        FontStyle Font { get; set; }
    }
}
