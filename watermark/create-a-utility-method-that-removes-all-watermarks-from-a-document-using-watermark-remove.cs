using System;
using System.IO;
using Aspose.Words;

public class Program
{
    // Removes any watermark present in the document.
    public static void RemoveAllWatermarks(Document doc)
    {
        if (doc.Watermark.Type != WatermarkType.None)
        {
            doc.Watermark.Remove();
        }
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a sample text watermark.
        doc.Watermark.SetText("Sample Watermark");

        // Save the document with the watermark (optional demonstration).
        string withWatermarkPath = "DocumentWithWatermark.docx";
        doc.Save(withWatermarkPath);

        // Remove all watermarks using the utility method.
        RemoveAllWatermarks(doc);

        // Save the document after watermark removal.
        string withoutWatermarkPath = "DocumentWithoutWatermark.docx";
        doc.Save(withoutWatermarkPath);
    }
}
