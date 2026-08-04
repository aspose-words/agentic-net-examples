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
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample document.");

        // Save the sample document to a memory stream.
        using (MemoryStream sourceStream = new MemoryStream())
        {
            sampleDoc.Save(sourceStream, SaveFormat.Docx);
            sourceStream.Position = 0; // Reset stream position for reading.

            // Load the document from the memory stream.
            Document doc = new Document(sourceStream);

            // Define watermark options (optional).
            TextWatermarkOptions options = new TextWatermarkOptions
            {
                FontFamily = "Arial",
                FontSize = 48,
                Color = Color.Red,
                Layout = WatermarkLayout.Diagonal,
                IsSemitrasparent = false
            };

            // Apply a text watermark.
            doc.Watermark.SetText("CONFIDENTIAL", options);

            // Save the watermarked document to disk.
            doc.Save("Watermarked.docx");
        }
    }
}
