using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Extract3DModelThumbnails
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourceDocx = @"C:\Docs\SourceDocument.docx";

        // Folder where extracted PNG thumbnails will be saved.
        string outputFolder = @"C:\Docs\Thumbnails";
        Directory.CreateDirectory(outputFolder);

        // Load the DOCX document (lifecycle rule: load).
        Document doc = new Document(sourceDocx);

        // Retrieve all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int thumbnailIndex = 0;

        // Iterate through each shape and extract its image data if present.
        foreach (Shape shape in shapeNodes)
        {
            // 3‑D model objects store a preview image (thumbnail) as the shape's image.
            // The HasImage property indicates that the shape contains image data.
            if (shape.HasImage)
            {
                // Build a unique file name for each extracted thumbnail.
                string pngPath = Path.Combine(outputFolder, $"Thumbnail_{thumbnailIndex}.png");

                // Save the image data as PNG (lifecycle rule: save).
                // The file extension determines the output format.
                shape.ImageData.Save(pngPath);

                thumbnailIndex++;
            }
        }

        Console.WriteLine($"Extracted {thumbnailIndex} thumbnail image(s) to \"{outputFolder}\".");
    }
}
