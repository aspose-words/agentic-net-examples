using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Settings;

class InsertImageWatermark
{
    static void Main()
    {
        // Path to the image that will be used as a watermark.
        const string imagePath = @"C:\Images\Logo.jpg";

        // Path where the resulting document will be saved.
        const string outputPath = @"C:\Output\DocumentWithImageWatermark.docx";

        // Create a new blank Word document.
        Document doc = new Document();

        // Optimize the document for a specific Microsoft Word version (e.g., Word 2016).
        // This prevents the document from opening in Compatibility mode.
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2016);

        // Configure optional image watermark settings.
        ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
        {
            // Scale factor (0 = auto). Adjust as needed.
            Scale = 0,
            // Set to false if you want the watermark to appear solid instead of washed out.
            IsWashout = true
        };

        // Add the image watermark to the document using the file path.
        doc.Watermark.SetImage(imagePath, watermarkOptions);

        // Save the document to the specified file.
        doc.Save(outputPath);
    }
}
