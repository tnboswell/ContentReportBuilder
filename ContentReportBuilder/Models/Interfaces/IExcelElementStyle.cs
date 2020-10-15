
namespace ContentReportBuilder.Models.Styles
{
    /// <summary>
    /// Element style for excel document elements. This interface adds to document base element styles.
    /// </summary>
    public interface IExcelElementStyle: IDocumentElementStyle
    {
        CellStyle Cell { get; set; }
    }
}
