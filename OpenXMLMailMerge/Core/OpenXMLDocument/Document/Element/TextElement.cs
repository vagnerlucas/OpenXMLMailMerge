using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace OpenXMLMailMerge.Core.OpenXMLDocument.Document.Element
{
    /// <summary>
    /// Text element class.
    /// </summary>
    internal static class TextElement
    {
        /// <summary>
        /// Creates a text element.
        /// </summary>
        /// <param name="text">Text value.</param>
        /// <returns>Text OpenXmlElement.</returns>
        internal static OpenXmlElement CreateText(string text)
        {
            if (text != null)
                return new Text() { Text = text };

            throw new ArgumentNullException("Invalid argument");
        }
    }
}
