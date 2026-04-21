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
        string watermarkText = "Sample Watermark";
        doc.Watermark.SetText(watermarkText);

        // Verify that the watermark was added.
        bool watermarkAdded = doc.Watermark.Type == WatermarkType.Text;
        Console.WriteLine($"Watermark added: {watermarkAdded}");

        // Remove the watermark.
        doc.Watermark.Remove();

        // Verify that the watermark was removed.
        bool watermarkRemoved = doc.Watermark.Type == WatermarkType.None;
        Console.WriteLine($"Watermark removed: {watermarkRemoved}");

        // Save the document to the local file system (optional verification of output file existence).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "WatermarkRemovalTest.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
