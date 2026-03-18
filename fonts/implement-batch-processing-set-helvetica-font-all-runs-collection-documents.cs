using System;
using System.IO;
using Aspose.Words;

public class HelveticaBatchProcessor
{
    /// <summary>
    /// Sets the font of every Run in each document to Helvetica and saves the modified document.
    /// </summary>
    /// <param name="inputFiles">Array of full paths to the source documents.</param>
    /// <param name="outputFolder">Folder where the processed documents will be saved.</param>
    public static void Process(string[] inputFiles, string outputFolder)
    {
        // Ensure the output directory exists.
        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        foreach (string inputPath in inputFiles)
        {
            // Load the document using the standard Aspose.Words constructor.
            Document doc = new Document(inputPath);

            // Iterate over all Run nodes in the document tree.
            foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
            {
                // Apply Helvetica to the run's font.
                run.Font.Name = "Helvetica";
            }

            // Build the output file name (preserve original name).
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(inputPath));

            // Save the modified document using the standard Save method.
            doc.Save(outputPath);
        }
    }
}

public class Program
{
    /// <summary>
    /// Entry point for the console application.
    /// Usage: dotnet run <inputFolder> <outputFolder>
    /// </summary>
    public static void Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: <inputFolder> <outputFolder>");
            return;
        }

        string inputFolder = args[0];
        string outputFolder = args[1];

        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Get all supported document files (e.g., .docx, .doc, .rtf, .txt, .pdf, etc.)
        string[] supportedExtensions = new[] { ".docx", ".doc", ".rtf", ".txt", ".pdf" };
        var inputFiles = Directory.GetFiles(inputFolder)
                                   .Where(f => supportedExtensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase))
                                   .ToArray();

        if (inputFiles.Length == 0)
        {
            Console.WriteLine($"No supported documents found in {inputFolder}");
            return;
        }

        HelveticaBatchProcessor.Process(inputFiles, outputFolder);
        Console.WriteLine($"Processed {inputFiles.Length} document(s). Output saved to {outputFolder}");
    }
}
