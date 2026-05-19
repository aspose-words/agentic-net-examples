using System;
using System.IO;
using Aspose.Words;

public class Program
{
    // Standardized disclaimer text to be added as a comment.
    private const string DisclaimerText = "Disclaimer: This document is confidential and intended solely for the recipient.";

    public static void Main()
    {
        // Prepare input and output folders relative to the current working directory.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample DOCX files for the batch process.
        CreateSampleDocuments(inputDir);

        // Process each DOCX file in the input folder.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Add the disclaimer comment.
            AddDisclaimerComment(doc);

            // Determine the output path (same file name, different folder).
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));

            // Save the modified document.
            doc.Save(outputPath);
        }
    }

    // Generates simple documents to demonstrate the batch operation.
    private static void CreateSampleDocuments(string folder)
    {
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample document #{i}");
            builder.Writeln("This is some example content.");
            string fileName = Path.Combine(folder, $"Sample{i}.docx");
            doc.Save(fileName);
        }
    }

    // Appends a standardized disclaimer comment to the last paragraph of the document.
    private static void AddDisclaimerComment(Document doc)
    {
        // Ensure the document has at least one paragraph to attach the comment to.
        Paragraph? targetParagraph = doc.LastSection?.Body?.LastParagraph;
        if (targetParagraph == null)
        {
            // If no paragraph exists, create a new one in the first section.
            doc.EnsureMinimum();
            targetParagraph = doc.FirstSection.Body.FirstParagraph;
        }

        // Create the comment node.
        Comment comment = new Comment(doc)
        {
            Author = "System",
            Initial = "SYS",
            DateTime = DateTime.Now
        };

        // Build the comment's visible content (Paragraph + Run).
        Paragraph commentParagraph = new Paragraph(doc);
        commentParagraph.AppendChild(new Run(doc, DisclaimerText));
        comment.AppendChild(commentParagraph);

        // Append the comment to the target paragraph.
        targetParagraph.AppendChild(comment);
    }
}
