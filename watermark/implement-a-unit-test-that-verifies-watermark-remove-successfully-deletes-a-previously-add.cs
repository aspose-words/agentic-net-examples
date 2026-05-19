using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a text watermark.
        const string watermarkText = "Test Watermark";
        doc.Watermark.SetText(watermarkText);

        // Verify that the watermark was added.
        bool watermarkAdded = doc.Watermark.Type == WatermarkType.Text;
        Console.WriteLine($"Watermark added: {watermarkAdded}");

        // Save the document with the watermark (optional, just to have an output file).
        string addedPath = Path.Combine(Directory.GetCurrentDirectory(), "WatermarkAdded.docx");
        doc.Save(addedPath);

        // Remove the watermark.
        doc.Watermark.Remove();

        // Verify that the watermark was removed.
        bool watermarkRemoved = doc.Watermark.Type == WatermarkType.None;
        Console.WriteLine($"Watermark removed: {watermarkRemoved}");

        // Save the document after removal.
        string removedPath = Path.Combine(Directory.GetCurrentDirectory(), "WatermarkRemoved.docx");
        doc.Save(removedPath);

        // Simple test result output.
        if (watermarkAdded && watermarkRemoved && File.Exists(addedPath) && File.Exists(removedPath))
        {
            Console.WriteLine("Test passed.");
        }
        else
        {
            Console.WriteLine("Test failed.");
        }
    }
}
