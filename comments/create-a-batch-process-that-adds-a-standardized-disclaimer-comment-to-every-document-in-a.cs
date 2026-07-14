using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare input and output folders in the current working directory.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample DOCX files to demonstrate the batch process.
        for (int i = 1; i <= 3; i++)
        {
            string samplePath = Path.Combine(inputDir, $"Sample{i}.docx");
            Document sampleDoc = new Document();
            DocumentBuilder sampleBuilder = new DocumentBuilder(sampleDoc);
            sampleBuilder.Writeln($"This is the content of sample document {i}.");
            sampleDoc.Save(samplePath);
        }

        // Standardized disclaimer comment text.
        const string disclaimerText = "Standard disclaimer: This document is confidential.";

        // Process each DOCX file in the input directory.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Insert the disclaimer comment at the end of the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();
            builder.Writeln(); // Ensure there is a paragraph to attach the comment to.

            // Create the comment node.
            Comment disclaimerComment = new Comment(doc)
            {
                Author = "Standard Disclaimer",
                Initial = "SD",
                DateTime = DateTime.Now
            };
            // Set the comment body text (creates necessary paragraphs and runs).
            disclaimerComment.SetText(disclaimerText);

            // Attach the comment to the current (last) paragraph.
            builder.CurrentParagraph?.AppendChild(disclaimerComment);

            // Save the modified document to the output folder, preserving the original file name.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }

        // The batch process completes here. No user interaction is required.
    }
}
