using System;
using System.IO;
using Aspose.Words;

public class Program
{
    // Entry point of the console application.
    public static void Main()
    {
        // Define the collection of document file paths to process.
        // In a real scenario, populate this array with actual file locations.
        string[] documentPaths = new string[]
        {
            "Document1.docx",
            "Document2.docx",
            "Document3.docx"
        };

        // Process each document.
        foreach (string inputPath in documentPaths)
        {
            // Verify the input file exists before attempting to load.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                continue;
            }

            // Load the document using Aspose.Words.
            Document doc = new Document(inputPath);

            // Iterate over all Run nodes in the document and set the font to Helvetica.
            foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
            {
                // Use the Font property of the Run node to set the font name.
                run.Font.Name = "Helvetica";
            }

            // Define the output file name (original name with "_out" suffix).
            string outputPath = Path.Combine(
                Path.GetDirectoryName(inputPath) ?? string.Empty,
                Path.GetFileNameWithoutExtension(inputPath) + "_out" + Path.GetExtension(inputPath));

            // Save the modified document.
            doc.Save(outputPath);

            // Validate that the output file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"Successfully processed and saved: {outputPath}");
            }
            else
            {
                Console.WriteLine($"Failed to save the document: {outputPath}");
            }
        }
    }
}
