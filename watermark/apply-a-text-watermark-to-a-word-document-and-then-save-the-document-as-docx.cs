using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");

        // Apply a text watermark.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Save the watermarked document.
        string outputPath = "Watermarked.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Watermarked document saved successfully to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the watermarked document.");
        }
    }
}
