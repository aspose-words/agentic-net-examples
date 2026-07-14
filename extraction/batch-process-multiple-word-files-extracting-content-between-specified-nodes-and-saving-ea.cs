using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "ExtractedPdfs");

        if (Directory.Exists(inputDir))
            Directory.Delete(inputDir, true);
        if (Directory.Exists(outputDir))
            Directory.Delete(outputDir, true);

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample documents with bookmarks
        const int sampleCount = 3;
        for (int i = 1; i <= sampleCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {i} - Intro paragraph.");
            builder.StartBookmark("Start");
            builder.Writeln($"Document {i} - Content paragraph 1.");
            builder.Writeln($"Document {i} - Content paragraph 2.");
            builder.EndBookmark("Start");
            builder.StartBookmark("End");
            builder.Writeln($"Document {i} - Outro paragraph.");
            builder.EndBookmark("End");

            string filePath = Path.Combine(inputDir, $"Sample{i}.docx");
            doc.Save(filePath);
        }

        // Process each document: extract content between the two bookmarks and save as PDF
        string[] files = Directory.GetFiles(inputDir, "*.docx");
        foreach (string file in files)
        {
            // Load source document
            Document sourceDoc = new Document(file);

            // Locate bookmarks
            Bookmark startBookmark = sourceDoc.Range.Bookmarks["Start"];
            Bookmark endBookmark = sourceDoc.Range.Bookmarks["End"];
            if (startBookmark == null || endBookmark == null)
                throw new InvalidOperationException($"Bookmarks not found in {Path.GetFileName(file)}.");

            // Extract the text that lies between the bookmarks.
            // The Bookmark.Text property returns the concatenated text of all nodes inside the bookmark.
            string extractedText = startBookmark.Text?.Trim();
            if (string.IsNullOrEmpty(extractedText))
                extractedText = "[No content between bookmarks]";

            // Create a new document to hold the extracted content.
            Document extractedDoc = new Document();
            DocumentBuilder ebuilder = new DocumentBuilder(extractedDoc);
            ebuilder.Writeln(extractedText);

            // Save the extracted document as PDF
            string outputFileName = Path.GetFileNameWithoutExtension(file) + "_Extracted.pdf";
            string outputPath = Path.Combine(outputDir, outputFileName);
            extractedDoc.Save(outputPath, SaveFormat.Pdf);

            // Validate that the PDF was created
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create PDF for {Path.GetFileName(file)}.");
        }

        // All done
        Console.WriteLine("Batch extraction completed successfully.");
    }
}
