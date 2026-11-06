using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a tiny PNG image (1x1 pixel) and save it locally.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
        const string imageFileName = "sample.png";
        File.WriteAllBytes(imageFileName, pngBytes);

        // -----------------------------------------------------------------
        // Step 1: Create a sample document that contains a picture content control.
        // -----------------------------------------------------------------
        Document inputDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(inputDoc);

        builder.Writeln("Document with a picture content control:");

        // Create an inline picture content control.
        StructuredDocumentTag pictureSdt = new StructuredDocumentTag(inputDoc, SdtType.Picture, MarkupLevel.Inline);
        // Insert the image into the content control.
        Shape pictureShape = new Shape(inputDoc, ShapeType.Image);
        pictureShape.ImageData.SetImage(imageFileName);
        pictureSdt.AppendChild(pictureShape);

        // Append the content control to the current paragraph.
        Paragraph currentParagraph = builder.CurrentParagraph;
        currentParagraph.AppendChild(pictureSdt);

        builder.Writeln("End of document.");

        const string inputPath = "input.docx";
        inputDoc.Save(inputPath);

        // -----------------------------------------------------------------
        // Step 2: Load the document and replace all picture content controls with inline images.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);

        // Find all picture content controls.
        List<StructuredDocumentTag> pictureControls = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.Picture)
            .ToList();

        foreach (StructuredDocumentTag sdt in pictureControls)
        {
            // Locate the image (Shape) inside the picture content control.
            Shape? shape = sdt.GetChildNodes(NodeType.Shape, true)
                .OfType<Shape>()
                .FirstOrDefault();

            if (shape == null)
                continue; // No image found; skip this control.

            // Extract the image bytes.
            using MemoryStream imageStream = new MemoryStream();
            shape.ImageData.Save(imageStream);
            imageStream.Position = 0;

            // Insert the image inline at the position of the picture content control.
            DocumentBuilder insertBuilder = new DocumentBuilder(doc);
            insertBuilder.MoveTo(sdt);
            insertBuilder.InsertImage(imageStream);

            // Remove the original picture content control.
            sdt.Remove();
        }

        // -----------------------------------------------------------------
        // Step 3: Save the resulting document.
        // -----------------------------------------------------------------
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
