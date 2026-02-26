using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCM document from disk.
        Document doc = new Document("Input.docm");

        // Remove the watermark if one is present (text or image).
        if (doc.Watermark.Type != WatermarkType.None)
        {
            doc.Watermark.Remove();
        }

        // Save the modified document back to disk.
        doc.Save("Output.docm");
    }
}
