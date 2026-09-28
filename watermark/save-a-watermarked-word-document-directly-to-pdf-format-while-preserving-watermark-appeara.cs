using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a paragraph with sample text so the PDF is not empty.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document with a text watermark.");

        // Apply a text watermark to the document.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Define the output PDF file path.
        string outputPath = "Watermarked.pdf";

        // Save the document directly to PDF format.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the PDF file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the PDF file.");
        }
    }
}
