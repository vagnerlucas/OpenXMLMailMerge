namespace OpenXmlMailMerge
{
    /// <summary>
    /// Type of data to merge fields.
    /// </summary>
    public enum MailMergeDataTypeEnum
    {
        /// <summary>
        /// Text is a string that replaces the field. The value is inserted as is.
        /// </summary>
        Text,
        /// <summary>
        /// Image is an array of byte to include within the document.
        /// </summary>
        Image,
        /// <summary>
        /// Table is a System.DataTable to include within the document.
        /// </summary>
        Table
    }
}
