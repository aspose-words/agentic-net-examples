using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample document in memory.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document created for watermark demonstration.");

        // Save the sample document to a memory stream.
        using (MemoryStream sourceStream = new MemoryStream())
        {
            sourceDoc.Save(sourceStream, SaveFormat.Docx);
            sourceStream.Position = 0; // Reset stream position for reading.

            // Load the document from the memory stream.
            Document doc = new Document(sourceStream);

            // Define watermark options (optional).
            TextWatermarkOptions options = new TextWatermarkOptions
            {
                FontFamily = "Arial",
                FontSize = 36,
                Color = Color.Gray,
                Layout = WatermarkLayout.Diagonal,
                IsSemitrasparent = false
            };

            // Apply a text watermark.
            doc.Watermark.SetText("Sample Watermark", options);

            // Save the watermarked document to disk.
            string outputPath = "Watermarked.docx";
            doc.Save(outputPath, SaveFormat.Docx);

            // Simple verification that the file was created.
            Console.WriteLine(File.Exists(outputPath)
                ? $"Watermarked document saved to '{outputPath}'."
                : "Failed to save the watermarked document.");
        }
    }
}
