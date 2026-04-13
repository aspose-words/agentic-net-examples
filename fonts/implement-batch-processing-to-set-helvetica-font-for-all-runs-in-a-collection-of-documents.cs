using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define folders for input and output documents.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample documents to work with.
        CreateSampleDocuments(inputDir);

        // Get all .docx files from the input directory.
        string[] inputFiles = Directory.GetFiles(inputDir, "*.docx");

        // Process each document: set all Run fonts to Helvetica and save.
        foreach (string inputFile in inputFiles)
        {
            // Load the document.
            Document doc = new Document(inputFile);

            // Retrieve all Run nodes in the document.
            NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);

            // Set the font name for each Run.
            foreach (Run run in runs)
            {
                run.Font.Name = "Helvetica";

                // Validate that the font name was set correctly.
                if (!string.Equals(run.Font.Name, "Helvetica", StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException($"Failed to set font for run: {run.GetText()}");
                }
            }

            // Determine the output file path.
            string outputFile = Path.Combine(outputDir, Path.GetFileName(inputFile));

            // Save the modified document.
            doc.Save(outputFile);

            // Verify that the output file exists.
            if (!File.Exists(outputFile))
            {
                throw new FileNotFoundException($"The document was not saved correctly: {outputFile}");
            }
        }

        // Optional: indicate processing is complete.
        Console.WriteLine("Batch font update completed successfully.");
    }

    // Helper method to create sample documents with some text.
    private static void CreateSampleDocuments(string folderPath)
    {
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample document {i}");
            builder.Writeln("This is a line of text.");
            builder.Writeln("Another line with different content.");

            string filePath = Path.Combine(folderPath, $"Doc{i}.docx");
            doc.Save(filePath);
        }
    }
}
