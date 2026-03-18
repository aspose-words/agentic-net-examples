using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Saving;
using Aspose.Words.Rendering;

class ExtractOfficeMathImages
{
    static void Main()
    {
        // Path to the source Word document containing OfficeMath objects.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Folder where the rendered PNG images will be saved.
        string outputFolder = @"C:\Docs\OfficeMathImages";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the document.
        Document doc = new Document(inputPath);

        // Retrieve all OfficeMath nodes in the document.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath object and render it to a PNG file.
        int index = 0;
        foreach (OfficeMath math in mathNodes)
        {
            // Configure image save options – PNG format.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png);

            // Optional: increase the scale for higher resolution output.
            // saveOptions.Scale = 5f; // Uncomment to render at 5x size.

            // Build the output file name.
            string outputPath = Path.Combine(outputFolder, $"OfficeMath_{index:D4}.png");

            // Render the OfficeMath object to an image file.
            math.GetMathRenderer().Save(outputPath, saveOptions);

            index++;
        }

        Console.WriteLine($"Extracted {index} OfficeMath images to \"{outputFolder}\".");
    }
}
