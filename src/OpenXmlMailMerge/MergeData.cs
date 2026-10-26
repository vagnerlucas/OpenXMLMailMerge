#nullable enable

using System;
using System.Collections.Generic;
using OpenXmlMailMerge.Merging;

namespace OpenXmlMailMerge
{
    /// <summary>
    /// The data to merge into a template: named text, image and table values.
    /// A name matches a field whose name is exactly the same, ignoring case.
    /// Build it once and merge it into any number of templates: a merge reads it without changing it.
    /// Adding values while a merge reads the same instance is not supported.
    /// </summary>
    public sealed class MergeData
    {
        private readonly Dictionary<string, (BindingKind Kind, object? Value)> _bindings = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<string> _names = [];

        /// <summary>
        /// The binding names, in the order they were added.
        /// </summary>
        public IReadOnlyList<string> Names
        {
            get { return _names; }
        }

        /// <summary>
        /// Binds a value to the text fields with this name. Numbers and dates are formatted in the template's language;
        /// null clears the field.
        /// </summary>
        /// <param name="name">The field name.</param>
        /// <param name="value">The value.</param>
        /// <returns>This instance, to chain calls.</returns>
        /// <exception cref="ArgumentException">The name is empty, or already bound.</exception>
        public MergeData Text(string name, object? value)
        {
            return Add(name, BindingKind.Text, value);
        }

        /// <summary>
        /// Binds an image to the image fields with this name. The bytes are copied, so later changes to the caller's
        /// buffer don't affect the merge.
        /// </summary>
        /// <param name="name">The field name.</param>
        /// <param name="image">A PNG, JPEG, GIF, BMP, TIFF, EMF or WMF file.</param>
        /// <returns>This instance, to chain calls.</returns>
        /// <exception cref="ArgumentException">The name is empty or already bound, or the image is empty.</exception>
        public MergeData Image(string name, ReadOnlyMemory<byte> image)
        {
            if (image.IsEmpty)
                throw new ArgumentException("The image is empty.", nameof(image));

            return Add(name, BindingKind.Image, image.ToArray());
        }

        /// <summary>
        /// Binds a table to the table fields with this name.
        /// </summary>
        /// <param name="name">The field name.</param>
        /// <param name="table">The table.</param>
        /// <returns>This instance, to chain calls.</returns>
        /// <exception cref="ArgumentNullException">The table is null.</exception>
        /// <exception cref="ArgumentException">The name is empty, or already bound.</exception>
        public MergeData Table(string name, TableValue table)
        {
            ArgumentNullException.ThrowIfNull(table);

            return Add(name, BindingKind.Table, table);
        }

        /// <summary>
        /// Gets the binding with this name, ignoring case.
        /// </summary>
        /// <param name="name">The binding name.</param>
        /// <param name="kind">The kind of value.</param>
        /// <param name="value">The value.</param>
        /// <returns>True when a value with this name is bound.</returns>
        internal bool TryGetBinding(string name, out BindingKind kind, out object? value)
        {
            if (_bindings.TryGetValue(name, out var binding))
            {
                (kind, value) = binding;
                return true;
            }

            kind = default;
            value = null;
            return false;
        }

        private MergeData Add(string name, BindingKind kind, object? value)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(name);

            if (!_bindings.TryAdd(name, (kind, value)))
                throw new ArgumentException($"A value named \"{name}\" is already bound. Names ignore case.", nameof(name));

            _names.Add(name);
            return this;
        }
    }
}
