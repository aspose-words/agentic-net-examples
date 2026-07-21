using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;
using Newtonsoft.Json; // Required by category rules

namespace RemovePictureContentControls
{
    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a sample document that contains a picture content control.
            // -----------------------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Document with a picture content control:");

            // Create an inline picture content control.
            StructuredDocumentTag pictureSdt = new StructuredDocumentTag(doc, SdtType.Picture, MarkupLevel.Inline)
            {
                Title = "SamplePicture",
                Tag = "pic1"
            };

            // Insert the content control into the first paragraph.
            Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
            firstParagraph.AppendChild(pictureSdt);

            // Tiny 1x1 PNG image (Base64 encoded).
            const string pngBase64 =
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            byte[] imageBytes = Convert.FromBase64String(pngBase64);

            // Create a Shape that holds the image and add it to the content control.
            Shape pictureShape = new Shape(doc, ShapeType.Image);
            // ImageData.SetImage does not accept a byte[] directly; use a MemoryStream.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                pictureShape.ImageData.SetImage(ms);
            }
            pictureShape.Width = 50;   // points
            pictureShape.Height = 50;  // points
            pictureSdt.AppendChild(pictureShape);

            // Add a blank paragraph after the control.
            builder.Writeln();
            builder.Writeln("End of sample document.");

            // Save the original document.
            const string inputPath = "InputWithPictureContentControl.docx";
            doc.Save(inputPath);

            // -----------------------------------------------------------------
            // 2. Remove all picture content controls while keeping their images.
            // -----------------------------------------------------------------
            // Find all StructuredDocumentTag nodes of type Picture.
            NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
            List<StructuredDocumentTag> pictureControls = new List<StructuredDocumentTag>();

            foreach (StructuredDocumentTag sdt in sdtNodes.OfType<StructuredDocumentTag>())
            {
                if (sdt.SdtType == SdtType.Picture)
                    pictureControls.Add(sdt);
            }

            // Remove each picture content control but retain its child (the image).
            foreach (StructuredDocumentTag pictureControl in pictureControls)
            {
                pictureControl.RemoveSelfOnly();
            }

            // Save the modified document.
            const string outputPath = "OutputWithoutPictureContentControls.docx";
            doc.Save(outputPath);
        }
    }
}
