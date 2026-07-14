using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Settings; // Needed for MsWordVersion

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // Create a simple PNG image to be used as a watermark.
        string imagePath = Path.Combine(dataDir, "watermark.png");
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII=";
        File.WriteAllBytes(imagePath, Convert.FromBase64String(base64Png));

        // Build a large document by inserting many paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 2000; i++)
        {
            builder.Writeln($"This is paragraph number {i + 1}.");
        }

        // Optimize the document for a specific Word version to improve processing speed.
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

        // Configure image watermark options.
        ImageWatermarkOptions imgOptions = new ImageWatermarkOptions
        {
            Scale = 0.5,          // Scale the watermark to 50% of the page width.
            IsWashout = false    // Keep the original colors of the image.
        };

        // Apply the image watermark using the file path.
        doc.Watermark.SetImage(imagePath, imgOptions);

        // Save the document with memory optimization enabled.
        string outputPath = Path.Combine(dataDir, "OptimizedWatermarked.docx");
        SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Docx);
        saveOptions.MemoryOptimization = true;
        doc.Save(outputPath, saveOptions);
    }
}
