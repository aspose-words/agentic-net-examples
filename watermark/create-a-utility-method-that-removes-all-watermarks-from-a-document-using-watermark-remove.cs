using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();

        // Add a text watermark to demonstrate removal.
        doc.Watermark.SetText("Sample Watermark");

        // Save the document with the watermark.
        doc.Save("WithWatermark.docx");

        // Remove all watermarks using the utility method.
        RemoveAllWatermarks(doc);

        // Save the document after watermark removal.
        doc.Save("WithoutWatermark.docx");
    }

    // Utility method that removes any existing watermark from the document.
    public static void RemoveAllWatermarks(Document doc)
    {
        // Check if a watermark is present and remove it.
        if (doc.Watermark.Type != WatermarkType.None)
        {
            doc.Watermark.Remove();
        }
    }
}
