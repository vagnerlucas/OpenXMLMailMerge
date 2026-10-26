using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OpenXmlMailMerge.Merging
{
    /// <summary>
    /// Hands out the ids new drawings need, so they never clash with the template's and the same merge always produces
    /// the same ids.
    /// </summary>
    internal sealed class IdAllocator
    {
        private uint _nextDrawingId;

        /// <summary>
        /// Starts drawing ids after the highest wp:docPr id in the given parts. Word expects these ids to be unique, so the
        /// parts should be every one that can hold drawings: the body, headers, footers, footnotes, endnotes and comments.
        /// </summary>
        /// <param name="roots">The root element of each part.</param>
        internal IdAllocator(IEnumerable<OpenXmlPartRootElement> roots)
        {
            var highest = roots.SelectMany(root => root.Descendants<DW.DocProperties>())
                .Select(properties => properties.Id?.Value ?? 0U)
                .DefaultIfEmpty(0U)
                .Max();

            _nextDrawingId = highest + 1;
        }

        /// <summary>
        /// Gets the next drawing id (wp:docPr id).
        /// </summary>
        /// <returns>An id no other drawing in the document uses.</returns>
        internal uint NextDrawingId()
        {
            return _nextDrawingId++;
        }

        /// <summary>
        /// Gets the first relationship id of the form "rIdN" that a part doesn't use yet.
        /// </summary>
        /// <param name="part">The part the new relationship starts from.</param>
        /// <returns>The relationship id.</returns>
        internal static string NextRelationshipId(OpenXmlPart part)
        {
            var used = new HashSet<string>(
                part.Parts.Select(pair => pair.RelationshipId)
                    .Concat(part.ExternalRelationships.Select(relationship => relationship.Id))
                    .Concat(part.HyperlinkRelationships.Select(relationship => relationship.Id))
                    .Concat(part.DataPartReferenceRelationships.Select(relationship => relationship.Id)),
                StringComparer.Ordinal);

            for (var i = 1; ; i++)
            {
                var id = $"rId{i}";
                if (!used.Contains(id))
                    return id;
            }
        }
    }
}
