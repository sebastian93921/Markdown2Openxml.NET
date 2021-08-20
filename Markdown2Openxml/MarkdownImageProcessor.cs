using System;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Net.Http;
using System.Net;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text.RegularExpressions;
using Markdown2Openxml.Enumeration;

namespace Markdown2Openxml
{
    public class MarkdownImageProcessor
    {
        public MarkdownImageProcessor()
        {
        }

        //Set Image width
        private const long AvailableWidth = 7600;

        private const double EmuPerPixel = 9525;
        private const long DocumentSizePerPixel = 15;
        private static uint docId = 1;

        public static Drawing convertMarkdownImageToRunElement(MainDocumentPart mainDocumentPart, string markdownString)
        {
            string imageBase64String = Regex.Replace(markdownString, @"\!\[(.+?)\]\(", "");
            if (imageBase64String.StartsWith("data:image")){

                imageBase64String = imageBase64String.Remove(imageBase64String.Length - 1, 1);
                string[] splitImageB64String = imageBase64String.Split(";");
                string datatype = splitImageB64String[0];
                string imageString = splitImageB64String[1].Replace("base64,","");
                byte[] imageByte = System.Convert.FromBase64String(imageString);

                ImagePartType imagePartType;
                switch (datatype)
                {
                    case "data:image/png":
                        imagePartType = ImagePartType.Png;
                        break;
                    case "data:image/jpg":
                    case "data:image/jpeg":
                        imagePartType = ImagePartType.Jpeg;
                        break;
                    default:
                        // Unsupported datatype
                        return null;
                }

                ImagePart imagePart = mainDocumentPart.AddImagePart(imagePartType);
                MemoryStream M = new MemoryStream(imageByte);
                imagePart.FeedData(M);

                ImageSize imageSize = determineSize(imagePartType, new MemoryStream(imageByte));

                long imageX = 990000L;
                long imageY = 792000L;
                if(imageSize != null){
                    Console.WriteLine("Original Image Size: "+imageSize.Width+"/"+imageSize.Height);
                    // Image actual size in px
                    double imageRatio = (double)imageSize.Height / imageSize.Width;

                    // Resize (Convert actual px to Emus)
                    imageX = (long)(EmuPerPixel * (AvailableWidth / DocumentSizePerPixel));
                    imageY = (long)(imageRatio * imageX);
                    Console.WriteLine("Calculated Image Size: "+imageX+"/"+imageY);
                }

                // Define the reference of the image.
                Drawing element =
                    new Drawing(
                        new DW.Inline(
                            new DW.Extent() { Cx = imageX, Cy = imageY },
                            new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, 
                                RightEdge = 0L, BottomEdge = 0L },
                            new DW.DocProperties() { Id = docId++, 
                                Name = "Image"+docId },
                            new DW.NonVisualGraphicFrameDrawingProperties(
                                new A.GraphicFrameLocks() { NoChangeAspect = true }),
                            new A.Graphic(
                                new A.GraphicData(
                                    new PIC.Picture(
                                        new PIC.NonVisualPictureProperties(
                                            new PIC.NonVisualDrawingProperties() 
                                            { Id = docId, 
                                                Name = "Image"+docId },
                                            new PIC.NonVisualPictureDrawingProperties()),
                                        new PIC.BlipFill(
                                            new A.Blip(){
                                                Embed = mainDocumentPart.GetIdOfPart(imagePart), 
                                                CompressionState = 
                                                A.BlipCompressionValues.Print },
                                            new A.Stretch(
                                                new A.FillRectangle())),
                                        new PIC.ShapeProperties(
                                            new A.Transform2D(
                                                new A.Offset() { X = 0L, Y = 0L },
                                                new A.Extents() { Cx = imageX, Cy = imageY }),
                                            new A.PresetGeometry(
                                                new A.AdjustValueList()
                                            ) { Preset = A.ShapeTypeValues.Rectangle }))
                                ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                        ) { DistanceFromTop = (UInt32Value)0U, 
                            DistanceFromBottom = (UInt32Value)0U, 
                            DistanceFromLeft = (UInt32Value)0U, 
                            DistanceFromRight = (UInt32Value)0U });
                return element;
            }
            return null;
        }


        private static ImageSize determineSize(ImagePartType imagePartType, Stream imageStream){
            using (BinaryReader reader = new BinaryReader(imageStream, System.Text.Encoding.ASCII, true))
            {
                switch(imagePartType){
                    case ImagePartType.Jpeg:
                        // JPEG magic number
                        var magicNumber = reader.ReadByte() << 8 | reader.ReadByte(); 

                        do
                        {
                            // Find next segment marker. Markers are zero or more 0xFF bytes, followed
                            // by a 0xFF and then a byte not equal to 0x00 or 0xFF.
                            byte segmentIdentifier = reader.ReadByte();
                            byte segmentType = reader.ReadByte();

                            // Read until we have a 0xFF byte followed by a byte that is not 0xFF or 0x00
                            while (segmentIdentifier != 0xFF || segmentType == 0xFF || segmentType == 0)
                            {
                                segmentIdentifier = segmentType;
                                segmentType = reader.ReadByte();
                            }

                            if (segmentType == 0xD9) // EOF?
                                return null;

                            // next 2-bytes are <segment-size>: [high-byte] [low-byte]
                            int segmentLength = reader.ReadByte() << 8 | reader.ReadByte();

                            // segment length includes size bytes, so subtract two
                            segmentLength -= 2;

                            if (segmentType == 0xC0 || segmentType == 0xC2)
                            {
                                reader.ReadByte();
                                // bits/sample, usually 8
                                int jpgheight = reader.ReadByte() << 8 | reader.ReadByte();
                                int jpgwidth = reader.ReadByte() << 8 | reader.ReadByte();
                                return new ImageSize(jpgwidth, jpgheight);
                            }
                            else
                            {
                                // skip this segment
                                reader.ReadBytes(segmentLength);
                            }
                        }
                        while (true);
                    case ImagePartType.Png:
                        byte[] pngSignatureBytes = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
                        reader.ReadBytes(pngSignatureBytes.Length);
                        reader.ReadBytes(8);
                        // Read 2 int32 as there png size
                        int pngwidth = reader.ReadByte() << 24 | reader.ReadByte() << 16 | reader.ReadByte() << 8  | reader.ReadByte();
                        int pngheight = reader.ReadByte() << 24 | reader.ReadByte() << 16 | reader.ReadByte() << 8  | reader.ReadByte();
                        return new ImageSize(pngwidth, pngheight);
                    default:
                        return null;
                }
            }
        }
    }
}
