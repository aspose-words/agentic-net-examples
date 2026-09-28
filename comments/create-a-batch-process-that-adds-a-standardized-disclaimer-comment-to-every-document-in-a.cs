using System;
using System.IO;
using System.Linq;
using Aspose.Words;

#nullable enable

public class BatchCommentAdder
{
    // Standardized disclaimer text to add as a comment.
    private const string DisclaimerText = "Disclaimer: This document is confidential.";

    // Author metadata for the disclaimer comment.
    private const string DisclaimerAuthor = "Compliance Team";
    private const string DisclaimerInitial = "CT";

    public static void Main()
    {
        // Define input and output directories relative to the current working directory.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "input");
        string outputDir = Path.Combine(baseDir, "output");

        // Ensure clean directories.
        if (Directory.Exists(inputDir))
            Directory.Delete(inputDir, true);
        if (Directory.Exists(outputDir))
            Directory.Delete(outputDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample documents to demonstrate the batch process.
        CreateSampleDocuments(inputDir);

        // Process each .docx file in the input directory.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Add the standardized disclaimer comment.
            AddDisclaimerComment(doc);

            // Determine output path and save the modified document.
            string fileName = Path.GetFileName(filePath);
            string outputPath = Path.Combine(outputDir, fileName);
            doc.Save(outputPath);
        }

        // The program finishes automatically; no user interaction required.
    }

    // Generates a few simple Word documents for the example.
    private static void CreateSampleDocuments(string folder)
    {
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document #{i}.");
            builder.Writeln("It contains some placeholder text for testing.");

            string filePath = Path.Combine(folder, $"Sample{i}.docx");
            doc.Save(filePath);
        }
    }

    // Adds a disclaimer comment to the first paragraph of the document.
    private static void AddDisclaimerComment(Document doc)
    {
        // Ensure the document has at least one section, body, and paragraph.
        Section? firstSection = doc.FirstSection;
        if (firstSection == null) return;

        Body? body = firstSection.Body;
        if (body == null) return;

        Paragraph? targetParagraph = body.FirstParagraph;
        if (targetParagraph == null)
        {
            // If no paragraph exists, create one.
            targetParagraph = new Paragraph(doc);
            body.AppendChild(targetParagraph);
        }

        // Create the comment node with metadata.
        Comment disclaimerComment = new Comment(doc)
        {
            Author = DisclaimerAuthor,
            Initial = DisclaimerInitial,
            DateTime = DateTime.Now
        };

        // Add a paragraph and run inside the comment to hold the disclaimer text.
        Paragraph commentParagraph = new Paragraph(doc);
        Run commentRun = new Run(doc, DisclaimerText);
        commentParagraph.AppendChild(commentRun);
        disclaimerComment.AppendChild(commentParagraph);

        // Anchor the comment to the target paragraph.
        // In Aspose.Words, a comment can be added as a child of the paragraph.
        targetParagraph.AppendChild(disclaimerComment);
    }
}
