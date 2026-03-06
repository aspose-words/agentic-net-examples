using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCM document from disk.
        Document doc = new Document("input.docm");

        // If the document contains any watermark (text or image), remove it.
        if (doc.Watermark.Type != WatermarkType.None)
        {
            doc.Watermark.Remove();
        }

        // Save the modified document back to disk (watermark removed).
        doc.Save("output.docm");
    }
}
