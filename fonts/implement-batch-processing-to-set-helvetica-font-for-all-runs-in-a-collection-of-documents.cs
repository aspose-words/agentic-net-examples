using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define folders for source and processed documents.
        string baseDir = Directory.GetCurrentDirectory();
        string sourceDir = Path.Combine(baseDir, "SourceDocs");
        string outputDir = Path.Combine(baseDir, "ProcessedDocs");

        Directory.CreateDirectory(sourceDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample documents.
        List<string> sourceFiles = new List<string>();
        for (int i = 1; i <= 3; i++)
        {
            string filePath = Path.Combine(sourceDir, $"Doc{i}.docx");
            CreateSampleDocument(filePath, $"This is sample document {i}.");
            sourceFiles.Add(filePath);
        }

        // Process each document: set all Run fonts to Helvetica.
        foreach (string srcPath in sourceFiles)
        {
            // Load the document.
            Document doc = new Document(srcPath);

            // Iterate over all Run nodes and set the font name.
            NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
            foreach (Run run in runs)
            {
                run.Font.Name = "Helvetica";
                // Simple validation that the font name was set.
                if (run.Font.Name != "Helvetica")
                {
                    throw new InvalidOperationException("Failed to set font to Helvetica.");
                }
            }

            // Save the modified document.
            string fileName = Path.GetFileNameWithoutExtension(srcPath);
            string outPath = Path.Combine(outputDir, $"{fileName}_Processed.docx");
            doc.Save(outPath);
        }

        // Optional: confirm that processed files exist.
        foreach (string file in Directory.GetFiles(outputDir, "*.docx"))
        {
            Console.WriteLine($"Processed file created: {file}");
        }
    }

    // Helper method to create a simple document with some text.
    private static void CreateSampleDocument(string path, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(path);
    }
}
