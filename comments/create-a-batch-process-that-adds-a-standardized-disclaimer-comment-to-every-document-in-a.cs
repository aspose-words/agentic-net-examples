using System;
using System.IO;
using Aspose.Words;

public class BatchAddDisclaimer
{
    public static void Main()
    {
        // Define input and output folders relative to the current directory.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");

        // Ensure a clean environment.
        if (Directory.Exists(inputDir))
            Directory.Delete(inputDir, true);
        if (Directory.Exists(outputDir))
            Directory.Delete(outputDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample documents for the batch process.
        CreateSampleDocument(Path.Combine(inputDir, "Sample1.docx"),
            "First sample document.");
        CreateSampleDocument(Path.Combine(inputDir, "Sample2.docx"),
            "Second sample document with more text.");

        // Standardized disclaimer text to be added as a comment.
        const string disclaimer = "This document is confidential. Do not distribute.";

        // Process each .docx file in the input folder.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Create a new comment node.
            Comment comment = new Comment(doc)
            {
                Author = "System",
                Initial = "SYS",
                DateTime = DateTime.Now
            };

            // Add a paragraph with the disclaimer text inside the comment.
            Paragraph commentParagraph = new Paragraph(doc);
            commentParagraph.AppendChild(new Run(doc, disclaimer));
            comment.AppendChild(commentParagraph);

            // Anchor the comment to the first paragraph of the document, if it exists.
            Paragraph? targetParagraph = doc.FirstSection?.Body?.FirstParagraph;
            if (targetParagraph != null)
            {
                targetParagraph.AppendChild(comment);
            }

            // Save the modified document to the output folder.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }
    }

    // Helper method that creates a simple one‑paragraph document with the supplied text.
    private static void CreateSampleDocument(string filePath, string bodyText)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(bodyText);
        doc.Save(filePath);
    }
}
