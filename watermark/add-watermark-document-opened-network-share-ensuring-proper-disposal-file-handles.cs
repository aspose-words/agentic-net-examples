using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;   // Required for watermark related types

class WatermarkExample
{
    static void Main()
    {
        // UNC path to the source document on a network share.
        string sourcePath = @"\\ServerName\ShareFolder\SourceDocument.docx";

        // Path where the watermarked document will be saved.
        string outputPath = @"C:\Temp\WatermarkedDocument.docx";

        // Open the document via a FileStream inside a using block to ensure the file handle is released.
        using (FileStream stream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            // Load the document from the stream (uses the provided Document(Stream) rule).
            Document doc = new Document(stream);

            // Configure text watermark options.
            TextWatermarkOptions options = new TextWatermarkOptions
            {
                FontFamily = "Arial",
                FontSize = 48,
                Color = System.Drawing.Color.Red,
                Layout = WatermarkLayout.Diagonal,
                IsSemitrasparent = false   // Fully opaque.
            };

            // Apply the watermark (uses the Document.Watermark.SetText method).
            doc.Watermark.SetText("CONFIDENTIAL", options);

            // Save the modified document (uses the Document.Save(string) rule).
            doc.Save(outputPath);
        }

        // At this point the FileStream has been disposed and the network file handle is released.
        Console.WriteLine("Watermark added and document saved to: " + outputPath);
    }
}
