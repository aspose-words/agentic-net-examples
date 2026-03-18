using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractAudioThumbnails
{
    static void Main()
    {
        // Load the Word document.
        Document doc = new Document("InputDocument.docx");

        // Get all shapes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int thumbnailIndex = 0;

        // Iterate through each shape.
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            // Check if the shape contains an image (e.g., an audio icon).
            if (shape.HasImage)
            {
                // Retrieve the raw image bytes from the shape.
                byte[] imageBytes = shape.ImageData.ImageBytes;

                // Define the output JPEG file name.
                string outputPath = $"AudioThumbnail_{thumbnailIndex}.jpg";

                // Write the bytes directly to a JPEG file. The image stored in the shape is already a raster image
                // (usually PNG or JPEG). If you need to force a specific format, you would have to load it with a
                // graphics library, but that would re‑introduce the System.Drawing dependency.
                File.WriteAllBytes(outputPath, imageBytes);

                thumbnailIndex++;
            }
        }

        Console.WriteLine($"Extracted {thumbnailIndex} thumbnail(s).");
    }
}
