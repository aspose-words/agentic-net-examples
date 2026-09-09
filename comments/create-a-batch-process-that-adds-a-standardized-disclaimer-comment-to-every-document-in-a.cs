using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    // Standardized disclaimer text to be added as a comment.
    private const string DisclaimerText = "Disclaimer: This document is confidential and intended for the designated recipient only.";

    // Author metadata for the disclaimer comment.
    private const string DisclaimerAuthor = "Compliance Team";
    private const string DisclaimerInitial = "CT";

    public static void Main()
    {
        // Define input and output folders relative to the current working directory.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // If the input folder is empty, create a few sample documents for demonstration.
        if (!Directory.EnumerateFiles(inputFolder, "*.docx").Any())
        {
            CreateSampleDocument(Path.Combine(inputFolder, "Sample1.docx"), "First sample document.");
            CreateSampleDocument(Path.Combine(inputFolder, "Sample2.docx"), "Second sample document with multiple paragraphs.\nSecond line.\nThird line.");
        }

        // Process each .docx file in the input folder.
        foreach (string inputPath in Directory.EnumerateFiles(inputFolder, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(inputPath);

            // Ensure the document has at least one paragraph to attach the comment.
            doc.EnsureMinimum();

            // Add the disclaimer comment to the last paragraph of the document.
            AddDisclaimerComment(doc);

            // Determine the output file path (same file name, different folder).
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(inputPath));

            // Save the modified document.
            doc.Save(outputPath);
        }
    }

    // Creates a simple document with the specified text content.
    private static void CreateSampleDocument(string filePath, string content)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(content);
        doc.Save(filePath);
    }

    // Adds a standardized disclaimer comment to the last paragraph of the given document.
    private static void AddDisclaimerComment(Document doc)
    {
        // Retrieve the last paragraph in the main body.
        var lastParagraph = doc.FirstSection?.Body?.LastParagraph;
        if (lastParagraph == null)
            return; // Safety check; should not occur because EnsureMinimum was called.

        // Create a new comment node.
        Comment comment = new Comment(doc)
        {
            Author = DisclaimerAuthor,
            Initial = DisclaimerInitial,
            DateTime = DateTime.Now
        };

        // Build the comment's visible content: a paragraph containing a run with the disclaimer text.
        comment.AppendChild(new Paragraph(doc));
        comment.FirstParagraph?.AppendChild(new Run(doc, DisclaimerText));

        // Append the comment to the paragraph so it appears in the document.
        lastParagraph.AppendChild(comment);
    }
}
