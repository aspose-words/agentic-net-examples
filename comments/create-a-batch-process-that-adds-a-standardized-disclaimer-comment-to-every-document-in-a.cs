using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare input and output folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample documents.
        CreateSampleDocuments(inputDir);

        // Process each document: add a standardized disclaimer comment.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Create the disclaimer comment.
            Comment disclaimer = new Comment(doc, "Standard Disclaimer", "SD", DateTime.Now);
            disclaimer.SetText("This document is confidential and intended solely for the recipient.");

            // Append the comment to the first paragraph of the document.
            Paragraph? firstParagraph = doc.FirstSection?.Body?.FirstParagraph;
            if (firstParagraph != null)
            {
                firstParagraph.AppendChild(disclaimer);
            }

            // Save the modified document to the output folder, preserving the original file name.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }

        // Indicate completion.
        Console.WriteLine("Disclaimer comments added to all documents.");
    }

    private static void CreateSampleDocuments(string folderPath)
    {
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample document {i}");
            builder.Writeln("This is some example content.");

            string fileName = Path.Combine(folderPath, $"Doc{i}.docx");
            doc.Save(fileName);
        }
    }
}
