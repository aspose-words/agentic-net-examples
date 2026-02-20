using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Customize the horizontal rule's formatting.
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // Centered on the page.
        format.WidthPercent = 80;                         // 80% of the page width.
        format.Height = 2;                                // 2 points thick.
        format.Color = Color.Black;                       // Black color.
        format.NoShade = true;                            // Solid color, no 3‑D shading.

        // Set up JPEG save options (you can adjust quality, resolution, etc.).
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            JpegQuality = 90 // High quality JPEG.
        };

        // Save the document as a JPEG image.
        doc.Save("HorizontalRule.jpg", jpegOptions);
    }
}
