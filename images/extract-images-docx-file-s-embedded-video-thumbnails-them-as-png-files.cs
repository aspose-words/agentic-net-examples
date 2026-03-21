using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractVideoThumbnails
{
    static void Main()
    {
        // Path to the DOCX that contains online video objects.
        // Use a relative path so the example works without requiring a specific absolute location.
        string inputFile = Path.Combine(AppContext.BaseDirectory, "VideoDocument.docx");

        // Folder where extracted PNG thumbnails will be saved.
        string outputFolder = Path.Combine(AppContext.BaseDirectory, "VideoThumbnails");
        Directory.CreateDirectory(outputFolder);

        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file not found: {inputFile}");
            Console.WriteLine("Place a DOCX file named 'VideoDocument.docx' in the application directory and rerun the program.");
            return;
        }

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

            // Build the output file name. The thumbnail image stored by Aspose.Words is already a PNG,
            // so we can save it directly.
            string outPath = Path.Combine(outputFolder, $"VideoThumbnail_{thumbIndex}.png");

            // Save the image data to a file.
            shape.ImageData.Save(outPath);
            thumbIndex++;
        }

        Console.WriteLine($"Extracted {thumbIndex} thumbnail(s) to \"{outputFolder}\".");
    }
}
