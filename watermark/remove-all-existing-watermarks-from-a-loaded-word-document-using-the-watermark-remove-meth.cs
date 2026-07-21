using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        const string inputPath = "SampleWithWatermark.docx";
        const string outputPath = "SampleWithoutWatermark.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a sample document and add a text watermark.
        // -----------------------------------------------------------------
        Document docWithWatermark = new Document();
        docWithWatermark.Watermark.SetText("Sample Watermark");
        docWithWatermark.Save(inputPath);

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
