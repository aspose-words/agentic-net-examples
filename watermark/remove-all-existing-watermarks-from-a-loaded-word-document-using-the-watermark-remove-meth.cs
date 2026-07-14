using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string dataDir = Path.Combine(Environment.CurrentDirectory, "Data");
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(dataDir);
        Directory.CreateDirectory(outputDir);

        // Path for the sample document that will contain a watermark.
        string samplePath = Path.Combine(dataDir, "Sample.docx");

        // Create a blank document and add a text watermark.
        Document doc = new Document();
        doc.Watermark.SetText("Sample Watermark");
        doc.Save(samplePath);

        // Load the document that has the watermark.
        Document loadedDoc = new Document(samplePath);

        // Remove the watermark if it exists.
        if (loadedDoc.Watermark.Type != WatermarkType.None)
        {
            loadedDoc.Watermark.Remove();
        }

        // Save the document without the watermark.
        string resultPath = Path.Combine(outputDir, "NoWatermark.docx");
        loadedDoc.Save(resultPath);
    }
}
