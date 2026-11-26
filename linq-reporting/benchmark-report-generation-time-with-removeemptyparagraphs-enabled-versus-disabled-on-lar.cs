using System;
using System.Diagnostics;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a large document with many paragraphs, including empty ones.
        string templatePath = Path.Combine(outputDir, "LargeDocument.docx");
        CreateLargeDocument(templatePath, paragraphCount: 5000, emptyParagraphInterval: 10);

        // Benchmark with empty paragraph removal disabled.
        string savedWithoutRemoval = Path.Combine(outputDir, "SavedWithoutRemoval.docx");
        long timeWithoutRemoval = SaveDocument(templatePath, savedWithoutRemoval, removeEmptyParagraphs: false);

        // Benchmark with empty paragraph removal enabled.
        string savedWithRemoval = Path.Combine(outputDir, "SavedWithRemoval.docx");
        long timeWithRemoval = SaveDocument(templatePath, savedWithRemoval, removeEmptyParagraphs: true);

        // Output results.
        Console.WriteLine($"Save time without removing empty paragraphs: {timeWithoutRemoval} ms");
        Console.WriteLine($"Save time with removing empty paragraphs:    {timeWithRemoval} ms");
    }

    private static void CreateLargeDocument(string filePath, int paragraphCount, int emptyParagraphInterval)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        for (int i = 1; i <= paragraphCount; i++)
        {
            if (i % emptyParagraphInterval == 0)
            {
                // Insert an empty paragraph.
                builder.Writeln();
            }
            else
            {
                builder.Writeln($"Paragraph {i}: The quick brown fox jumps over the lazy dog.");
            }
        }

        doc.Save(filePath, SaveFormat.Docx);
    }

    private static long SaveDocument(string sourcePath, string destinationPath, bool removeEmptyParagraphs)
    {
        var doc = new Document(sourcePath);

        // Remove empty paragraphs only when requested.
        if (removeEmptyParagraphs)
        {
            RemoveEmptyParagraphs(doc);
        }

        var stopwatch = Stopwatch.StartNew();
        doc.Save(destinationPath, SaveFormat.Docx);
        stopwatch.Stop();

        return stopwatch.ElapsedMilliseconds;
    }

    // Custom implementation to remove empty paragraphs from a document.
    private static void RemoveEmptyParagraphs(Document doc)
    {
        // Iterate through all paragraphs in reverse order to safely remove them.
        for (int i = doc.FirstSection.Body.Paragraphs.Count - 1; i >= 0; i--)
        {
            var paragraph = doc.FirstSection.Body.Paragraphs[i];
            // Get the paragraph text without the paragraph break.
            string text = paragraph.GetText().Replace("\r", string.Empty).Replace("\a", string.Empty);
            if (string.IsNullOrWhiteSpace(text))
            {
                paragraph.Remove();
            }
        }
    }
}
