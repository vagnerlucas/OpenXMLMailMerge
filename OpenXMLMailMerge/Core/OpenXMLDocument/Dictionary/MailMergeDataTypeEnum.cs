using System;

namespace OpenXMLMailMerge.Core.OpenXMLDocument.Dictionary
{
    /// <summary>
    /// Type of data to merge fields.
    /// </summary>
    public enum MailMergeDataTypeEnum
    {
        /// <summary>
        /// Regex is a single string to be replaced (using regex pattern).
        /// </summary>
        Regex,
        /// <summary>
        /// Image is a array of byte to include within the document.
        /// </summary>
        Image,
        /// <summary>
        /// Table is a System.DataTable to include within the document.
        /// </summary>
        Table
    }

    /// <summary>
    /// Extensions to display the name of the type.
    /// </summary>
    internal static class MailMergeDataTypeEnumExtensions
    {
        internal static string Description(this MailMergeDataTypeEnum mailMergeDataType)
        {
            switch (mailMergeDataType)
            {
                case MailMergeDataTypeEnum.Regex:
                    return "regex";
                case MailMergeDataTypeEnum.Image:
                    return "image";
                case MailMergeDataTypeEnum.Table:
                    return "table";
                default:
                    throw new ArgumentOutOfRangeException("Unknown data type");
            }
        }
    }
}
