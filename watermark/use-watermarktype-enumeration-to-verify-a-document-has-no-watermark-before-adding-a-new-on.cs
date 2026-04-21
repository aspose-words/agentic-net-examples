using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a simple paragraph so the document is not empty.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document content.");

        // Verify that the document currently has no watermark.
        if (doc.Watermark.Type == WatermarkType.None)
        {
            // Since there is no existing watermark, add a new text watermark.
            doc.Watermark.SetText("CONFIDENTIAL");
        }

        // Save the resulting document.
        const string outputPath = "WatermarkedDocument.docx";
        doc.Save(outputPath);

        // Indicate completion (optional, not required for verification).
        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
