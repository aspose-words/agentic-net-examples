using System;
using System.IO;
using Aspose.Words;

class RemoveWatermarksExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a sample text watermark so we have something to remove.
        doc.Watermark.SetText("Sample Watermark");

        // Remove the watermark if it exists.
        if (doc.Watermark.Type != WatermarkType.None)
        {
            doc.Watermark.Remove();
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocument.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
