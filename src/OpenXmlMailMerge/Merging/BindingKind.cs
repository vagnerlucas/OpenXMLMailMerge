namespace OpenXmlMailMerge.Merging
{
    /// <summary>
    /// The kind of value bound to a field.
    /// </summary>
    internal enum BindingKind
    {
        /// <summary>
        /// A value written as text.
        /// </summary>
        Text,
        /// <summary>
        /// Image bytes, inserted as an inline picture.
        /// </summary>
        Image,
        /// <summary>
        /// A <see cref="TableValue"/>, inserted as a table.
        /// </summary>
        Table
    }
}
