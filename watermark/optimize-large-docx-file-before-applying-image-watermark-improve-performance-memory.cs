using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Settings;

class WatermarkProcessor
{
    /// <summary>
    /// Loads a DOCX, optimizes it for performance, adds an image watermark,
    /// and saves the result using memory‑optimized save options.
    /// </summary>
    /// <param name="inputPath">Full path to the source DOCX file.</param>
    /// <param name="outputPath">Full path where the processed DOCX will be saved.</param>
    /// <param name="imagePath">Full path to the image to be used as a watermark.</param>
    public static void AddImageWatermark(string inputPath, string outputPath, string imagePath)
    {
        // Load the existing document.
        Document doc = new Document(inputPath);

        // Optimize compatibility for a recent Word version.
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2016);

        // Remove unused styles and lists – reduces document size and memory footprint.
        doc.Cleanup();

        // Rebuild the page layout so that any subsequent rendering uses up‑to‑date layout data.
        doc.UpdatePageLayout();

        // Configure watermark appearance.
        ImageWatermarkOptions wmOptions = new ImageWatermarkOptions
        {
            Scale = 0.5,          // 50 % of the original image size.
            IsWashout = false    // Keep original colors (no washout effect).
        };

        // Apply the image watermark.
        doc.Watermark.SetImage(imagePath, wmOptions);

        // Prepare save options that enable memory optimization during the save operation.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            MemoryOptimization = true   // Reduces memory consumption at the cost of a slightly longer save time.
        };

        // Save the processed document.
        doc.Save(outputPath, saveOptions);
    }

    // Example usage.
    static void Main()
    {
        // Create temporary files for the demo.
        string tempDir = Path.Combine(Path.GetTempPath(), "WatermarkDemo");
        Directory.CreateDirectory(tempDir);

        string sourceDoc = Path.Combine(tempDir, "LargeDocument.docx");
        string watermarkImg = Path.Combine(tempDir, "Watermark.png");
        string resultDoc = Path.Combine(tempDir, "LargeDocument_Watermarked.docx");

        // Create a simple DOCX if it does not exist.
        if (!File.Exists(sourceDoc))
        {
            Document demoDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(demoDoc);
            builder.Writeln("This is a sample document used for watermark demonstration.");
            demoDoc.Save(sourceDoc);
        }

        // Create a simple PNG image if it does not exist.
        if (!File.Exists(watermarkImg))
        {
            // 1x1 transparent PNG (base64 encoded)
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(watermarkImg, pngBytes);
        }

        // Apply the watermark.
        AddImageWatermark(sourceDoc, resultDoc, watermarkImg);
        Console.WriteLine($"Watermark applied and document saved to: {resultDoc}");
    }
}
