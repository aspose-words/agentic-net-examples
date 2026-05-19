using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a text watermark to demonstrate removal.
        doc.Watermark.SetText("Sample Watermark");

        // Save the document with the watermark (optional verification).
        doc.Save("WithWatermark.docx");

        // Remove all watermarks from the document.
        RemoveAllWatermarks(doc);

        // Save the document after watermark removal.
        doc.Save("WithoutWatermark.docx");
    }

    // Removes any existing watermark from the provided document.
    public static void RemoveAllWatermarks(Document doc)
    {
        // If a watermark is present, remove it.
        if (doc.Watermark.Type != WatermarkType.None)
        {
            doc.Watermark.Remove();
        }
    }
}
