using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document content.");

        // Verify that the document currently has no watermark.
        if (doc.Watermark.Type == WatermarkType.None)
        {
            // Since no watermark exists, add a text watermark.
            doc.Watermark.SetText("CONFIDENTIAL");
        }

        // Save the resulting document.
        const string outputPath = "WatermarkedDocument.docx";
        doc.Save(outputPath);
    }
}
