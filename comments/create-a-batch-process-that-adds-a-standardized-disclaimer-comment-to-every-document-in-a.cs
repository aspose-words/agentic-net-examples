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

        // Create a few sample DOCX files to demonstrate the batch process.
        for (int i = 1; i <= 2; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample document {i} content.");
            string samplePath = Path.Combine(inputDir, $"Doc{i}.docx");
            sampleDoc.Save(samplePath);
        }

        // Standardized disclaimer text to be added as a comment.
        const string disclaimerText = "Disclaimer: This document is confidential.";

        // Process each DOCX file in the input directory.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Ensure there is at least one paragraph to attach the comment.
            Paragraph? targetParagraph = doc.FirstSection?.Body?.FirstParagraph;
            if (targetParagraph == null)
            {
                // If the document has no paragraphs, create one.
                targetParagraph = new Paragraph(doc);
                doc.FirstSection?.Body?.AppendChild(targetParagraph);
            }

            // Create a new comment with author metadata.
            Comment comment = new Comment(doc, "System", "SYS", DateTime.Now);
            // Set the comment text (this also creates the necessary paragraph inside the comment).
            comment.SetText(disclaimerText);
            // Append the comment to the chosen paragraph.
            targetParagraph.AppendChild(comment);

            // Save the modified document to the output folder, preserving the original file name.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }
    }
}
