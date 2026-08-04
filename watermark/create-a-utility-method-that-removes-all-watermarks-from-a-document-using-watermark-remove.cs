using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();

        // Add a text watermark.
        doc.Watermark.SetText("Sample Watermark");

        // Save the document that contains the watermark.
        doc.Save("WithWatermark.docx");

        // Remove all watermarks using the utility method.
        RemoveAllWatermarks(doc);

        // Save the document after watermark removal.
        doc.Save("WithoutWatermark.docx");
    }

    // Utility method that removes any watermark from the given document.
    public static void RemoveAllWatermarks(Document document)
    {
        // Watermark.Remove() works only when a watermark is present.
        if (document.Watermark.Type != WatermarkType.None)
        {
            document.Watermark.Remove();
        }
    }
}
