using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Folder to store sample and output documents
        string folder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(folder);

        // Create sample documents
        List<string> docPaths = new List<string>();
        for (int i = 1; i <= 3; i++)
        {
            string path = Path.Combine(folder, $"Sample{i}.docx");
            CreateSampleDocument(path, $"This is sample document {i}.");
            docPaths.Add(path);
        }

        // Process each document: set Helvetica font for all runs
        foreach (string inputPath in docPaths)
        {
            Document doc = new Document(inputPath);

            // Iterate through all Run nodes
            foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
            {
                // Set font name to Helvetica
                run.Font.Name = "Helvetica";

                // Validate assignment
                if (run.Font.Name != "Helvetica")
                {
                    throw new InvalidOperationException("Font assignment failed.");
                }
            }

            // Save output document
            string outputPath = Path.Combine(folder, Path.GetFileNameWithoutExtension(inputPath) + "_out.docx");
            doc.Save(outputPath);

            // Ensure output file exists
            if (!File.Exists(outputPath))
            {
                throw new FileNotFoundException("Output file was not created.", outputPath);
            }
        }
    }

    // Helper method to create a simple document with a single paragraph
    private static void CreateSampleDocument(string filePath, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(filePath);
    }
}
