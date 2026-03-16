using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Settings;

class WatermarkProcessor
{
    /// <summary>
    /// Loads a large DOCX, optimizes it for performance, adds an image watermark,
    /// and saves the result using memory‑optimized save options.
    /// </summary>
    /// <param name="inputPath">Full path to the source DOCX file.</param>
    /// <param name="outputPath">Full path where the processed DOCX will be saved.</param>
    /// <param name="imagePath">Full path to the image to be used as a watermark.</param>
    public static void AddImageWatermark(string inputPath, string outputPath, string imagePath)
    {
        // Load the existing document (uses the Document(string) constructor).
        Document doc = new Document(inputPath);

        // Optimize compatibility for a recent Word version to avoid "Compatibility mode".
        // This can improve layout processing speed.
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2016);

        // Remove unused styles and lists – reduces document size and memory footprint.
        doc.Cleanup();

        // Rebuild the page layout so that any subsequent rendering (e.g., saving) uses up‑to‑date layout data.
        doc.UpdatePageLayout();

        // Configure watermark appearance.
        ImageWatermarkOptions wmOptions = new ImageWatermarkOptions
        {
            Scale = 0.5,          // 50 % of the original image size.
            IsWashout = false    // Keep original colors (no washout effect).
        };

        // Apply the image watermark using the provided SetImage overload.
        doc.Watermark.SetImage(imagePath, wmOptions);

        // Prepare save options that enable memory optimization during the save operation.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            MemoryOptimization = true   // Reduces memory consumption at the cost of a slightly longer save time.
        };

        // Save the processed document (uses Document.Save(string, SaveOptions)).
        doc.Save(outputPath, saveOptions);
    }

    // Example usage.
    static void Main()
    {
        string sourceDoc = @"C:\Docs\LargeDocument.docx";
        string watermarkImg = @"C:\Images\Watermark.png";
        string resultDoc = @"C:\Docs\LargeDocument_Watermarked.docx";

        AddImageWatermark(sourceDoc, resultDoc, watermarkImg);
        Console.WriteLine("Watermark applied and document saved.");
    }
}
