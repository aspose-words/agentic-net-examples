using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();

        // Add a text watermark.
        const string watermarkText = "Sample Watermark";
        doc.Watermark.SetText(watermarkText);

        // Verify that the watermark was added.
        if (doc.Watermark.Type != WatermarkType.Text)
        {
            Console.WriteLine("Error: Watermark was not added.");
            Environment.Exit(1);
        }

        // Save the document with the watermark.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);
        string watermarkedPath = Path.Combine(outputDir, "watermarked.docx");
        doc.Save(watermarkedPath);

        // Remove the watermark.
        doc.Watermark.Remove();

        // Verify that the watermark was removed.
        if (doc.Watermark.Type != WatermarkType.None)
        {
            Console.WriteLine("Error: Watermark was not removed.");
            Environment.Exit(1);
        }

        // Save the document after removal.
        string removedPath = Path.Combine(outputDir, "removed.docx");
        doc.Save(removedPath);

        Console.WriteLine("Watermark removal test passed.");
    }
}
