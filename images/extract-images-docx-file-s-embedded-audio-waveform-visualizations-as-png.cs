using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Rendering;

class ExtractAudioWaveformImages
{
    static void Main()
    {
        // Path to the source DOCX file that contains audio waveform visualizations.
        string sourceDocPath = @"C:\Input\AudioDocument.docx";

        // Folder where extracted PNG images will be saved.
        string outputFolder = @"C:\Output\WaveformImages";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the Word document.
        Document doc = new Document(sourceDocPath);

        // Retrieve all Shape nodes in the document (including those inside headers/footers).
        var shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;

        // Iterate through each shape and extract those that contain an image.
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Render the shape to a PNG image.
                ShapeRenderer renderer = shape.GetShapeRenderer();

                // Configure image save options to produce PNG output.
                ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);

                // Build the output file name.
                string pngFilePath = Path.Combine(outputFolder, $"waveform_{imageIndex}.png");

                // Save the rendered image.
                renderer.Save(pngFilePath, pngOptions);

                imageIndex++;
            }
        }

        Console.WriteLine($"Extracted {imageIndex} image(s) to \"{outputFolder}\".");
    }
}
