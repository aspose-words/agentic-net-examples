using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    // Entry point of the console application.
    public static void Main()
    {
        // Define folders for input and output documents.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Collect all .docx files from the input folder.
        List<string> documentPaths = new List<string>(Directory.GetFiles(inputFolder, "*.docx"));

        // If there are no documents, create a sample one to demonstrate the process.
        if (documentPaths.Count == 0)
        {
            string samplePath = Path.Combine(inputFolder, "Sample.docx");
            CreateSampleDocument(samplePath);
            documentPaths.Add(samplePath);
        }

        // Process each document: set all Run fonts to Helvetica and save.
        foreach (string docPath in documentPaths)
        {
            // Load the document.
            Document doc = new Document(docPath);

            // Iterate over all Run nodes in the document.
            NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
            foreach (Run run in runs)
            {
                // Set the font name to Helvetica.
                run.Font.Name = "Helvetica";

                // Validate that the font name was set correctly.
                if (!string.Equals(run.Font.Name, "Helvetica", StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException($"Failed to set font for run in document '{docPath}'.");
                }
            }

            // Save the modified document to the output folder, preserving the original file name.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(docPath));
            doc.Save(outputPath);
        }

        // Indicate completion (no interactive prompts).
        Console.WriteLine($"Processed {documentPaths.Count} document(s). Output saved to '{outputFolder}'.");
    }

    // Helper method to create a simple sample document if none are present.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        builder.Writeln("It contains multiple runs with default fonts.");
        doc.Save(filePath);
    }
}
