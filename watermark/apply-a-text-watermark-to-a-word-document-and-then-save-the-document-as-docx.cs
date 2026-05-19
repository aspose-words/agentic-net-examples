using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Apply a text watermark.
        doc.Watermark.SetText("Confidential");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "WatermarkedDocument.docx");

        // Save the document as DOCX.
        doc.Save(outputPath, SaveFormat.Docx);

        // Simple validation: ensure the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Watermark applied and document saved successfully.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
