namespace OpenXMLMailMerge.Core
{
    /// <summary>
    /// String constant utils.
    /// </summary>
    public static class StringConst
    {
        /// <summary>
        /// Regex to identify a mergefield from xml node
        /// </summary>
        public const string MERGEFIELD_REGEX = "(?<=MERGEFIELD)(.*)(?=MERGEFORMAT)";

        /// <summary>
        /// MSWord inputs this char in mergefield.
        /// </summary>
        public const string SPECIAL_STR_L = "«";
        /// <summary>
        /// MSWord inputs this char in mergefield.
        /// </summary>
        public const string SPECIAL_STR_R = "»";
        /// <summary>
        /// Parameter delimiter.
        /// </summary>
        public const char ID_DELIMITER_CHAR = '@';
        /// <summary>
        /// Parameter assignment delimiter.
        /// </summary>
        public const char KEY_VALUE_DELIMITER = '=';
    }
}
