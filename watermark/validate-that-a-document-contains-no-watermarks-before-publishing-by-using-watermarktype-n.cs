using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for output files.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);

        // Path for the sample document.
        string samplePath = Path.Combine(artifactsDir, "sample.docx");

        // 1. Create a blank document.
        Document doc = new Document();

        // 2. Save the blank document (no watermark present).
        doc.Save(samplePath);

        // 3. Load the document to simulate a publishing workflow.
        Document loadedDoc = new Document(samplePath);

        // 4. Validate that the document contains no watermark.
        if (loadedDoc.Watermark.Type == WatermarkType.None)
        {
            Console.WriteLine("Validation passed: No watermark present.");
        }
        else
        {
            Console.WriteLine("Validation failed: Watermark detected.");
        }

        // 5. If validation passes, save the document as the final published file.
        if (loadedDoc.Watermark.Type == WatermarkType.None)
        {
            string publishPath = Path.Combine(artifactsDir, "published.docx");
            loadedDoc.Save(publishPath);
        }
    }
}
