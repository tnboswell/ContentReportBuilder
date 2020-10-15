
namespace ContentReportBuilder.Models.Styles
{
    /// <summary>
    /// Element style for word document elements. This interface adds to document base element styles.
    /// </summary>
    public interface IWordElementStyle: IDocumentElementStyle
    {
        IndentationStyle Indentation { get; set; }
        SpacingStyle Spacing { get; set; }
    }
}
