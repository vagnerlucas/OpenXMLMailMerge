using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OpenXmlMailMerge.Elements
{
    /// <summary>
    /// Image element class.
    /// </summary>
    internal class ImageElement
    {
        /// <summary>
        /// Creates a Drawing OpenXml element.
        /// </summary>
        /// <param name="data">Byte array of the image data.</param>
        /// <param name="documentPart">Parent element.</param>
        /// <param name="imageElementConfig">Image element definitions.</param>
        /// <returns>Generic OpenXmlElement.</returns>
        internal static OpenXmlElement CreateImage(byte[] data, OpenXmlPart documentPart, ImageElementConfig imageElementConfig = null)
        {
            if (documentPart == null)
                throw new ArgumentNullException("Invalid document");

            if (data == null || data.Length == 0)
                return null;

            if (imageElementConfig == null)
            {
                imageElementConfig = new ImageElementConfig { Height = ImageElementConfig.DefaultSize, Width = ImageElementConfig.DefaultSize };
            }

            imageElementConfig.ResolveSize(data);

            var lcx = imageElementConfig.Width * 9525L;
            var lcy = imageElementConfig.Height * 9525L;

            var stream = new MemoryStream();
            stream.Write(data, 0, data.Length);
            stream.Seek(0, SeekOrigin.Begin);

            var contentType = ImageContentType.Detect(data);
            var relationshipId = $"r{new Random().Next().ToString()}";
            var imagePart = documentPart.AddNewPart<ImagePart>(contentType, relationshipId);

            imagePart.FeedData(stream);

            var randId = new Random().Next();
            var randomName = randId.ToString();

            var imageElement = new Drawing(
                new DW.Inline(new DW.Extent { Cx = lcx, Cy = lcy }, //{ Cx = 990000L, Cy = 792000L },
                    new DW.EffectExtent
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DW.DocProperties
                    {
                        Id = Convert.ToUInt32(randId),
                        Name = $"Pic{randomName}"
                    },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties
                                    {
                                        Id = Convert.ToUInt32(new Random().Next()),
                                        Name = $"Img{randomName}.{ImageContentType.GetExtension(contentType)}"
                                    },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip(
                                        new A.BlipExtensionList(
                                            new A.BlipExtension { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" }
                                        )
                                    )
                                    {
                                        Embed = relationshipId,
                                        CompressionState = A.BlipCompressionValues.Print
                                    },
                                    new A.Stretch(
                                        new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset { X = 0L, Y = 0L },
                                        new A.Extents { Cx = lcx, Cy = lcy }), //{ Cx = 990000L, Cy = 792000L }),
                                    new A.PresetGeometry(
                                        new A.AdjustValueList()
                                    )
                                    { Preset = A.ShapeTypeValues.Rectangle }))
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                )
                {
                    DistanceFromTop = (UInt32Value)0U,
                    DistanceFromBottom = (UInt32Value)0U,
                    DistanceFromLeft = (UInt32Value)0U,
                    DistanceFromRight = (UInt32Value)0U
                });

            return imageElement;
        }
    }
}
