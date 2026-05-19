using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define a folder for the sample files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Paths for the input (with watermark) and output (without watermark) documents.
        string inputPath = Path.Combine(artifactsDir, "DocumentWithWatermark.docx");
        string outputPath = Path.Combine(artifactsDir, "DocumentWithoutWatermark.docx");

        // -----------------------------------------------------------------
        // Step 1: Create a sample document and add a text watermark.
        // -----------------------------------------------------------------
        Document doc = new Document();
        doc.Watermark.SetText("Sample Watermark");
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // Step 2: Load the document that contains the watermark.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // -----------------------------------------------------------------
        // Step 3: Remove the watermark if it exists.
        // -----------------------------------------------------------------
        if (loadedDoc.Watermark.Type != WatermarkType.None)
        {
            loadedDoc.Watermark.Remove();
        }

        // -----------------------------------------------------------------
        // Step 4: Save the document without the watermark.
        // -----------------------------------------------------------------
        loadedDoc.Save(outputPath);
    }
}
