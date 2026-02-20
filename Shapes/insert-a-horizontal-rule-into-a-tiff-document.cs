using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Customize the horizontal rule's appearance (optional).
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // center the rule
        format.WidthPercent = 80;                         // 80% of the page width
        format.Height = 2;                                // 2 points high
        format.Color = Color.Black;                       // black color
        format.NoShade = true;                            // solid color, no 3‑D shading

        // Prepare TIFF save options.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // Render each page as a separate frame in the TIFF file.
        tiffOptions.PageLayout = MultiPageLayout.TiffFrames();
        // Optional: set compression for the TIFF image.
        tiffOptions.TiffCompression = TiffCompression.Ccitt4;

        // Save the document as a TIFF image.
        doc.Save("HorizontalRule.tiff", tiffOptions);
    }
}
