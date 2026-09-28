using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        // Add a paragraph so the document has content.
        doc.FirstSection.Body.AppendChild(new Paragraph(doc));

        // Apply a text watermark.
        doc.Watermark.SetText("Confidential");

        // Save the document with the watermark.
        string withWatermarkPath = "WithWatermark.docx";
        doc.Save(withWatermarkPath);

        // Remove all watermarks from the document.
        RemoveAllWatermarks(doc);

        // Save the document after watermark removal.
        string withoutWatermarkPath = "WithoutWatermark.docx";
        doc.Save(withoutWatermarkPath);

        // Simple validation that the files were saved.
        Console.WriteLine(File.Exists(withWatermarkPath) ? "Watermark file saved." : "Failed to save watermark file.");
        Console.WriteLine(File.Exists(withoutWatermarkPath) ? "Watermark removed file saved." : "Failed to save removed file.");
    }

    // Utility method that removes all watermarks from a document.
    public static void RemoveAllWatermarks(Document document)
    {
        document.Watermark.Remove();
    }
}
