using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractVideoThumbnails
{
    static void Main()
    {
        // Path to the DOCX that contains online video objects.
        string inputFile = @"C:\Docs\VideoDocument.docx";

        // Folder where extracted PNG thumbnails will be saved.
        string outputFolder = @"C:\Docs\VideoThumbnails";
        Directory.CreateDirectory(outputFolder);

        // Load the document.
        Document doc = new Document(inputFile);

        // Retrieve all Shape nodes – video objects are stored as shapes with an image (the thumbnail).
        var shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int thumbIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Only shapes that actually contain an image are candidates for thumbnails.
            if (!shape.HasImage)
                continue;

            // Build the output file name.  The thumbnail image stored by Aspose.Words is already a PNG,
            // so we can save it directly without using System.Drawing.
            string outPath = Path.Combine(outputFolder, $"VideoThumbnail_{thumbIndex}.png");

            // Save the image data to a file.  Aspose.Words will keep the original image format; if the
            // source thumbnail is not PNG you can change the extension to the appropriate format.
            shape.ImageData.Save(outPath);

            thumbIndex++;
        }

        Console.WriteLine($"Extracted {thumbIndex} thumbnail(s) to \"{outputFolder}\".");
    }
}
