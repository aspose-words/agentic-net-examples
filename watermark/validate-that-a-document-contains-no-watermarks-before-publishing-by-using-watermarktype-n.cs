using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        const string filePath = "sample.docx";

        // Create a new blank document.
        Document doc = new Document();

        // Save the document to a local file.
        doc.Save(filePath);

        // Load the document from the file.
        Document loadedDoc = new Document(filePath);

        // Validate that the document contains no watermark.
        // The Watermark.Type property returns a WatermarkType enum defined in Aspose.Words.
        if (loadedDoc.Watermark.Type == Aspose.Words.WatermarkType.None)
        {
            Console.WriteLine("Validation passed: No watermark present.");
        }
        else
        {
            Console.WriteLine("Validation failed: Watermark detected.");
        }
    }
}
