using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class WatermarkExample
{
    static void Main()
    {
        // Assume we already have the Word document bytes in memory.
        byte[] docBytes = File.ReadAllBytes("InputDocument.docx"); // Replace with your source bytes.

        // Load the document from a memory stream (no disk I/O for loading).
        using (MemoryStream inputStream = new MemoryStream(docBytes))
        {
            Document doc = new Document(inputStream); // Load constructor from Stream.

            // Apply a plain text watermark to the whole document.
            doc.Watermark.SetText("Confidential");

            // Save the modified document back to a memory stream (no disk I/O for saving).
            using (MemoryStream outputStream = new MemoryStream())
            {
                doc.Save(outputStream, SaveFormat.Docx); // Save to stream with desired format.

                // At this point outputStream contains the watermarked document.
                // Example: retrieve the byte array for further processing or transmission.
                byte[] watermarkedBytes = outputStream.ToArray();

                // (Optional) Write to console the size of the resulting document.
                Console.WriteLine($"Watermarked document size: {watermarkedBytes.Length} bytes");
            }
        }
    }
}
