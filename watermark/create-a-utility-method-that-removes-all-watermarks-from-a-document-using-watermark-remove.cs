using System;
using System.IO;
using Aspose.Words;

namespace WatermarkRemovalDemo
{
    class Program
    {
        static void Main()
        {
            // Create a sample document.
            Document doc = new Document();

            // Add a text watermark.
            doc.Watermark.SetText("Confidential");

            // Save the document with the watermark.
            string withWatermarkPath = "WithWatermark.docx";
            doc.Save(withWatermarkPath);

            // Remove all watermarks using the utility method.
            RemoveAllWatermarks(doc);

            // Save the document after watermark removal.
            string withoutWatermarkPath = "WithoutWatermark.docx";
            doc.Save(withoutWatermarkPath);

            // Validate that the files were created.
            Console.WriteLine($"Document with watermark saved: {File.Exists(withWatermarkPath)}");
            Console.WriteLine($"Document without watermark saved: {File.Exists(withoutWatermarkPath)}");
        }

        // Utility method that removes all watermarks from the provided document.
        static void RemoveAllWatermarks(Document document)
        {
            // Document.Watermark.Remove clears any existing watermark.
            document.Watermark.Remove();
        }
    }
}
