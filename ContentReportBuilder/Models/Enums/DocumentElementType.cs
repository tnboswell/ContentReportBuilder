
namespace ContentReportBuilder.Models
{
    /// <summary>
    /// Builder Supports for Word:
    /// - Heading
    /// - Paragraph
    /// - List
    /// - Table
    /// - Image
    /// 
    /// Builder Supports Excel:
    /// - Heading
    /// - Paragraph
    /// - List
    /// - Table
    /// - Image
    /// - Line Chart
    /// 
    /// </summary>
    public enum DocumentElementType
    {
        Heading = 1,
        Paragraph = 2,
        List = 3,
        Table = 4,
        Image = 5,
        LineChart = 6,
        PieChart = 7,
        BarChart = 8,
        Header = 9,
        Footer = 10
    }
}
