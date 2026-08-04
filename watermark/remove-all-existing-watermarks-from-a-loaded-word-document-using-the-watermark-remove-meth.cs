using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define folder and file names.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string inputPath = Path.Combine(artifactsDir, "DocumentWithWatermark.docx");
        string outputPath = Path.Combine(artifactsDir, "DocumentWithoutWatermark.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document and add a text watermark.
        // -----------------------------------------------------------------
        Document docWithWatermark = new Document();
        docWithWatermark.Watermark.SetText("Sample Watermark");
        docWithWatermark.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document that contains the watermark.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Remove the watermark if it exists.
        // -----------------------------------------------------------------
        if (loadedDoc.Watermark.Type != WatermarkType.None)
        {
            loadedDoc.Watermark.Remove();
        }

        // -----------------------------------------------------------------
        // 4. Save the document without the watermark.
        // -----------------------------------------------------------------
        loadedDoc.Save(outputPath);
    }
}
