using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class ExtractVideoFrames
{
    static void Main()
    {
        // Path to the source DOCX that contains embedded online videos.
        // Use a path relative to the executable so the example can run without requiring a specific absolute location.
        string sourceDocPath = Path.Combine(Environment.CurrentDirectory, "VideoDocument.docx");

        // Folder where extracted high‑resolution PNG frames will be saved.
        string outputFolder = Path.Combine(Environment.CurrentDirectory, "ExtractedFrames");

        if (!File.Exists(sourceDocPath))
        {
            Console.WriteLine($"Source document not found: \"{sourceDocPath}\"");
            Console.WriteLine("Place a DOCX file named \"VideoDocument.docx\" in the program's working directory and rerun.");
            return;
        }

        // Load the document.
        Document doc = new Document(sourceDocPath);

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Collect all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int frameIndex = 0;

        // Iterate through each shape and extract the thumbnail image of video objects.
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Video objects are stored as shapes that contain an image (the thumbnail frame).
            // The HasImage property is true for such shapes.
            if (shape.HasImage)
            {
                // Configure high‑resolution PNG output.
                ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
                {
                    // 300 dpi provides a high‑quality raster image.
                    Resolution = 300
                };

                // Build the output file name.
                string outFile = Path.Combine(outputFolder, $"VideoFrame_{frameIndex}.png");

                // Render the shape (thumbnail) to a PNG file.
                shape.GetShapeRenderer().Save(outFile, pngOptions);

                frameIndex++;
            }
        }

        Console.WriteLine($"Extracted {frameIndex} video frame image(s) to \"{outputFolder}\".");
    }
}
