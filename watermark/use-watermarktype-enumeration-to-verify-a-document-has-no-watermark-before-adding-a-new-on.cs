using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define the output file path in the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "WatermarkedDocument.docx");

        // Create a new blank document.
        Document doc = new Document();

        // Verify that the document does not already contain a watermark.
        // If a watermark exists (type is not None), remove it to ensure a clean state.
        if (doc.Watermark.Type != WatermarkType.None)
        {
            doc.Watermark.Remove();
        }

        // Add a new text watermark because the document is confirmed to have no watermark.
        doc.Watermark.SetText("Confidential");

        // Save the document with the newly added watermark.
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
