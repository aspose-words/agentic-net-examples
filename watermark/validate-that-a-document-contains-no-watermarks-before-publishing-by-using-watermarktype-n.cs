using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a blank document and add some sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample content for watermark validation.");

        // Path for the output document.
        string outputPath = "ValidatedDocument.docx";

        // Validate that the document has no watermark.
        if (doc.Watermark.Type == WatermarkType.None)
        {
            Console.WriteLine("Validation passed: No watermark present.");
        }
        else
        {
            Console.WriteLine("Validation failed: Watermark detected. Removing it.");
            doc.Watermark.Remove();
        }

        // Save the (potentially cleaned) document.
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
