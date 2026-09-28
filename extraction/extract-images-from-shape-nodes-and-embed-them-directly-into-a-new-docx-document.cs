using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // Create a sample source document that contains an image inside a shape.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);

        // 1x1 pixel PNG image (base64 encoded).
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=");
        using (MemoryStream imageStream = new MemoryStream(pngBytes))
        {
            sourceBuilder.InsertImage(imageStream);
        }

        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -------------------------------------------------
        // Load the source document and locate all shapes that contain images.
        // -------------------------------------------------
        Document loadedDoc = new Document(sourcePath);
        var imageShapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                   .OfType<Shape>()
                                   .Where(s => s.HasImage)
                                   .ToList();

        if (imageShapes.Count == 0)
            throw new InvalidOperationException("No image-bearing shapes were found in the source document.");

        // -------------------------------------------------
        // Prepare an empty destination document.
        // -------------------------------------------------
        Document destDoc = new Document();
        destDoc.RemoveAllChildren(); // Ensure the document is empty.

        // Build the minimal required structure: Section -> Body.
        Section destSection = new Section(destDoc);
        destDoc.AppendChild(destSection);
        Body destBody = new Body(destDoc);
        destSection.AppendChild(destBody);

        // Use a DocumentBuilder positioned at the start of the body.
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.MoveToDocumentStart();

        // -------------------------------------------------
        // Extract each image from the source shape and embed it into the destination document.
        // -------------------------------------------------
        foreach (Shape shape in imageShapes)
        {
            using (MemoryStream extractedImage = new MemoryStream())
            {
                // Save the image data from the shape into a memory stream.
                shape.ImageData.Save(extractedImage);
                extractedImage.Position = 0; // Reset stream position before reading.

                // Insert the image into the destination document.
                destBuilder.InsertImage(extractedImage);
                // Add a paragraph break after each image for readability.
                destBuilder.Writeln();
            }
        }

        // -------------------------------------------------
        // Save the destination document containing the extracted images.
        // -------------------------------------------------
        const string destPath = "extracted-images.docx";
        destDoc.Save(destPath);

        // Verify that the output file was created.
        if (!File.Exists(destPath))
            throw new InvalidOperationException("The destination document was not created as expected.");
    }
}
