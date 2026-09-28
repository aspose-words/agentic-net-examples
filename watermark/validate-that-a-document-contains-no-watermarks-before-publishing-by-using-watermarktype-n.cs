using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document without any watermark.
        string docPath = "NoWatermark.docx";
        Document doc = new Document();
        doc.Save(docPath);

        // Load the document for validation.
        Document loadedDoc = new Document(docPath);

        // Check that the document has no watermark.
        // The WatermarkType enum is defined in the Aspose.Words namespace.
        bool hasNoWatermark = loadedDoc.Watermark.Type == Aspose.Words.WatermarkType.None;

        // Output the validation result.
        Console.WriteLine(hasNoWatermark ? "No watermark detected." : "Watermark detected.");

        // Write a simple validation report to a file.
        string reportPath = "validation.txt";
        File.WriteAllText(reportPath, hasNoWatermark ? "Document has no watermarks." : "Document contains watermarks.");
    }
}
