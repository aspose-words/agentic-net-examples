using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Verify that the document currently has no watermark.
        if (doc.Watermark.Type == WatermarkType.None)
        {
            // Add a text watermark because none exists.
            doc.Watermark.SetText("Sample Watermark");
        }

        // Define the output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");

        // Ensure the directory exists.
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

        // Save the document with the new watermark.
        doc.Save(outputPath);
    }
}
