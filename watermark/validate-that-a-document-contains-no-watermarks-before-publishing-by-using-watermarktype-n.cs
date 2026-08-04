using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string sourcePath = Path.Combine(artifactsDir, "Sample.docx");

        // Create a blank document (no watermark).
        Document doc = new Document();
        doc.EnsureMinimum();
        doc.Save(sourcePath);

        // Load the document to validate.
        Document loadedDoc = new Document(sourcePath);

        // Check that the document has no watermark.
        bool hasNoWatermark = loadedDoc.Watermark.Type == WatermarkType.None;
        Console.WriteLine(hasNoWatermark
            ? "Validation passed: No watermark present."
            : "Validation failed: Watermark detected.");

        // If validation succeeds, simulate publishing by saving a copy.
        if (hasNoWatermark)
        {
            string publishedPath = Path.Combine(artifactsDir, "Published.docx");
            loadedDoc.Save(publishedPath);
        }
    }
}
