namespace OpenXMLMailMerge.Core.OpenXMLDocument.Configuration
{
    /// <summary>
    /// Dynamic configuration set by user.
    /// </summary>
    public class GenericConfiguration
    {
        /// <summary>
        /// Dynamic property to be used as needed.
        /// The default object is a dictionary.
        /// </summary>
        public dynamic Configuration { get; set; }
    }
}
