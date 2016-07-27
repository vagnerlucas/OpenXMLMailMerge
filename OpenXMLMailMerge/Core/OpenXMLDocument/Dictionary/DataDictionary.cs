using System.Collections.Generic;

namespace OpenXMLMailMerge.Core.OpenXMLDocument.Dictionary
{
    /// <summary>
    /// This class define the document's dictionary.
    /// Once the client sets the type of data, the tag and its value, the processor works on these 
    /// data to merge fields.
    /// </summary>
    public class DataDictionary : Dictionary<MailMergeDataTypeEnum, Dictionary<string, object>> { }
}
