using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "WatermarkedDocument.docx");

        // Create a new blank document.
        Document doc = new Document();

        // (Optional) Add a paragraph so the document has visible content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document with a text watermark.");

        // Apply a text watermark to the document.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Save the document as DOCX.
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Watermark applied and document saved to: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the watermarked document.");
        }
    }
}
