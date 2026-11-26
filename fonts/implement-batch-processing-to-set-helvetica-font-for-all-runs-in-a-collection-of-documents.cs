using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define folders for input and output documents.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample documents if none exist in the input folder.
        if (Directory.GetFiles(inputFolder, "*.docx").Length == 0)
        {
            CreateSampleDocument(Path.Combine(inputFolder, "Sample1.docx"), "First sample document.");
            CreateSampleDocument(Path.Combine(inputFolder, "Sample2.docx"), "Second sample document.");
        }

        // Process each document: set Helvetica font for all runs.
        foreach (string inputPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(inputPath);

            // Iterate over all Run nodes and set the font.
            foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
            {
                run.Font.Name = "Helvetica";

                // Validation: ensure the font name was set correctly.
                if (!string.Equals(run.Font.Name, "Helvetica", StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException($"Failed to set font for run in document '{inputPath}'.");
                }
            }

            // Save the modified document to the output folder.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(inputPath));
            doc.Save(outputPath);

            // Verify that the output file was created.
            if (!File.Exists(outputPath))
            {
                throw new FileNotFoundException($"Output file was not created: {outputPath}");
            }
        }

        // Indicate completion (optional, not interactive).
        Console.WriteLine("Batch processing completed successfully.");
    }

    // Helper method to create a simple document with given text.
    private static void CreateSampleDocument(string filePath, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(filePath);
    }
}
