using System;
using Aspose.Words;

namespace WatermarkDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();

            // Verify that the document does not contain any watermark.
            // WatermarkType.None indicates the absence of a watermark.
            if (doc.Watermark.Type == WatermarkType.None)
            {
                // Add a text watermark because none exists.
                doc.Watermark.SetText("CONFIDENTIAL");
            }

            // Save the resulting document.
            const string outputPath = "Result.docx";
            doc.Save(outputPath);

            // Simple verification that the file was created.
            Console.WriteLine($"Document saved to '{outputPath}'. Watermark added: {doc.Watermark.Type != WatermarkType.None}");
        }
    }
}
