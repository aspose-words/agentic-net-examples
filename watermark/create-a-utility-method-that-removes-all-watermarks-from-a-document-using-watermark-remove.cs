using System;
using Aspose.Words;

namespace WatermarkRemovalDemo
{
    public class Program
    {
        // Removes any watermark (text or image) from the provided document.
        public static void RemoveAllWatermarks(Document doc)
        {
            // Watermark.Remove clears the watermark regardless of its type.
            if (doc.Watermark.Type != WatermarkType.None)
                doc.Watermark.Remove();
        }

        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Add a text watermark so we have something to remove.
            doc.Watermark.SetText("Sample Watermark");

            // Save the document with the watermark.
            doc.Save("Document_With_Watermark.docx");

            // Remove all watermarks using the utility method.
            RemoveAllWatermarks(doc);

            // Save the cleaned document.
            doc.Save("Document_Without_Watermark.docx");
        }
    }
}
