using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Check if the document currently has no watermark.
        if (doc.Watermark.Type == WatermarkType.None)
        {
            // Since there is no watermark, add a text watermark.
            doc.Watermark.SetText("Confidential");
        }

        // Save the resulting document to the local file system.
        doc.Save("Result.docx");
    }
}
