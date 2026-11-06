using System;
using Aspose.Words;

namespace WatermarkRemovalDemo
{
    class Program
    {
        static void Main()
        {
            // Create a sample document with some content.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample document.");

            // Apply a text watermark to demonstrate removal.
            doc.Watermark.SetText("CONFIDENTIAL");

            // Remove all watermarks using the utility method.
            RemoveAllWatermarks(doc);

            // Save the document without watermarks.
            string outputPath = "Output_NoWatermark.docx";
            doc.Save(outputPath);
        }

        // Utility method that removes all watermarks from the provided document.
        static void RemoveAllWatermarks(Document document)
        {
            // Watermark.Remove clears any watermark applied to the document.
            document.Watermark.Remove();
        }
    }
}
