using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Verify that the document does not already contain a watermark.
        if (doc.Watermark.Type == WatermarkType.None)
        {
            // Add a text watermark because none exists.
            doc.Watermark.SetText("CONFIDENTIAL");
        }

        // Save the document with the new watermark.
        string outputPath = "output.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        bool fileExists = File.Exists(outputPath);
        Console.WriteLine($"Document saved: {fileExists}");
    }
}
