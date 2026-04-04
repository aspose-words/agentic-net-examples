using System;
using System.IO;
using Aspose.Words;

public class Program
{
    // Standard disclaimer text to be added as a comment.
    private const string DisclaimerAuthor = "Standard Disclaimer";
    private const string DisclaimerInitial = "SD";
    private const string DisclaimerText = "This document is confidential and intended for the designated recipient only.";

    public static void Main()
    {
        // Prepare input and output folders inside the current working directory.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample DOCX files to demonstrate the batch process.
        CreateSampleDocuments(inputDir);

        // Process each DOCX file: load, add disclaimer comment, and save to output folder.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Locate the last paragraph in the document where the comment will be attached.
            Paragraph? lastParagraph = doc.LastSection?.Body?.LastParagraph;

            if (lastParagraph != null)
            {
                // Create a new comment with author, initials and current date/time.
                Comment disclaimerComment = new Comment(doc, DisclaimerAuthor, DisclaimerInitial, DateTime.Now);
                disclaimerComment.SetText(DisclaimerText);

                // Append the comment to the selected paragraph.
                lastParagraph.AppendChild(disclaimerComment);
            }

            // Save the modified document to the output directory, preserving the original file name.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }

        // Optional: indicate completion.
        Console.WriteLine("Batch processing completed. Modified files are located in: " + outputDir);
    }

    // Generates simple Word documents with placeholder text.
    private static void CreateSampleDocuments(string folderPath)
    {
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample document {i} content.");
            string fileName = Path.Combine(folderPath, $"Sample{i}.docx");
            doc.Save(fileName);
        }
    }
}
