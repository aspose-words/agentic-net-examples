using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add some sample content so the watermark can be seen.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document with a text watermark.");

        // Apply a text watermark using the native Watermark API.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Define the output path for the DOCX file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "WatermarkedDocument.docx");

        // Save the document as DOCX.
        doc.Save(outputPath);

        // Verify that the file was created and report the result.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Watermarked document saved successfully to: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the watermarked document.");
        }
    }
}
