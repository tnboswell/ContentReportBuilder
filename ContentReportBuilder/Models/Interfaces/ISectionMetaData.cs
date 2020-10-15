
namespace ContentReportBuilder.Models.MetaData
{
    /// <summary>
    /// Must be provided
    /// </summary>
    public interface ISectionMetaData
    {
        /// <value>
        /// For excel the is the sheet name
        /// For word this is the new section name
        /// </value>
        string Title { get; }
    }
}
