using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document without any watermark.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document without watermark.");

        // Save the document (optional, demonstrates output file creation).
        string outputPath = "NoWatermark.docx";
        doc.Save(outputPath);

        // Validate that the document contains no watermark.
        bool hasNoWatermark = doc.Watermark.Type == WatermarkType.None;

        // Report the validation result.
        Console.WriteLine(hasNoWatermark
            ? "Document has no watermark."
            : "Document contains a watermark.");

        // If a watermark were present, remove it and save a cleaned copy.
        if (!hasNoWatermark)
        {
            doc.Watermark.Remove();
            doc.Save("Cleaned_" + outputPath);
        }
    }
}
