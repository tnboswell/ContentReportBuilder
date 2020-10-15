

namespace ContentReportBuilder.Models.MetaData
{
    /// <summary>
    /// Must be provided
    /// </summary>
    public class SectionMetaData : ISectionMetaData
    {
        /// <summary>
        /// Constructor that must be used
        /// </summary>
        /// <param name="title"></param>
        public SectionMetaData(string title)
        {
            Title = title;
        }

        /// <value>
        /// For excel the is the sheet name
        /// For word this is the new section name
        /// </value>
        public string Title { get; protected set; }
    }
}
