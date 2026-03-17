using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Define optional settings for the image watermark.
        ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
        {
            Scale = 5,          // Scale factor for the watermark image.
            IsWashout = false   // Preserve original colors (no washout effect).
        };

        // Path to the image that will be used as the watermark.
        string imagePath = @"C:\Images\Logo.jpg";

        // Apply the image watermark to the document.
        doc.Watermark.SetImage(imagePath, watermarkOptions);

        // Save the watermarked document in DOCX format.
        doc.Save(@"C:\Output\WatermarkedDocument.docx");
    }
}
