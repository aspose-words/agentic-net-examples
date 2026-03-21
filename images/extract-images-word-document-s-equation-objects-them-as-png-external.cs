using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Saving;

class ExtractOfficeMathImages
{
    static void Main(string[] args)
    {
        // Allow optional command‑line arguments for input and output paths.
        string inputPath = args.Length > 0 ? args[0] : Path.Combine(Environment.CurrentDirectory, "SourceDocument.docx");
        string outputFolder = args.Length > 1 ? args[1] : Path.Combine(Environment.CurrentDirectory, "OfficeMathImages");

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // If the source document does not exist, create a minimal one (without OfficeMath) so the program can run.
        if (!File.Exists(inputPath))
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("This is a placeholder document. No OfficeMath objects are present.");
            doc.Save(inputPath);
            Console.WriteLine($"Created placeholder document at \"{inputPath}\".");
        }

        // Load the document.
        Document document = new Document(inputPath);

        // Retrieve all OfficeMath nodes in the document.
        NodeCollection mathNodes = document.GetChildNodes(NodeType.OfficeMath, true);

        int index = 0;
        foreach (OfficeMath math in mathNodes)
        {
            // Configure image save options – PNG format.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png);

            // Optional: increase the scale for higher resolution output.
            // saveOptions.Scale = 5f; // Uncomment to render at 5x size.

            string outputPath = Path.Combine(outputFolder, $"OfficeMath_{index:D4}.png");

            // Render the OfficeMath object to an image file.
            math.GetMathRenderer().Save(outputPath, saveOptions);
            index++;
        }

        Console.WriteLine($"Extracted {index} OfficeMath image(s) to \"{outputFolder}\".");
    }
}
