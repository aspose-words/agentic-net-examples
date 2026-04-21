using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the sample document and the result after removal.
        string sourcePath = Path.Combine(outputDir, "source.docx");
        string resultPath = Path.Combine(outputDir, "result.docx");

        // Create a blank document and add a text watermark.
        Document doc = new Document();
        doc.Watermark.SetText("Sample Watermark");
        doc.Save(sourcePath);

        // Load the document that contains the watermark.
        Document loadedDoc = new Document(sourcePath);

        // Remove the watermark if it exists.
        if (loadedDoc.Watermark.Type != WatermarkType.None)
        {
            loadedDoc.Watermark.Remove();
        }

        // Save the document without the watermark.
        loadedDoc.Save(resultPath);

        // Indicate completion.
        Console.WriteLine("Watermark removed. Output saved to: " + resultPath);
    }
}
