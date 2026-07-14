using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();

        // Add a text watermark.
        string watermarkText = "Test Watermark";
        doc.Watermark.SetText(watermarkText);

        // Verify that the watermark was added.
        bool added = doc.Watermark.Type == WatermarkType.Text;
        if (!added)
        {
            Console.WriteLine("Failed to add watermark.");
            return;
        }

        // Save the document with the watermark (optional, just to demonstrate saving).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Watermarked.docx");
        doc.Save(outputPath);

        // Remove the watermark.
        doc.Watermark.Remove();

        // Verify that the watermark was removed.
        bool removed = doc.Watermark.Type == WatermarkType.None;
        if (removed)
        {
            Console.WriteLine("Watermark removal test passed.");
        }
        else
        {
            Console.WriteLine("Watermark removal test failed.");
        }

        // Save the document after removal (optional).
        string cleanedPath = Path.Combine(Directory.GetCurrentDirectory(), "Cleaned.docx");
        doc.Save(cleanedPath);
    }
}
