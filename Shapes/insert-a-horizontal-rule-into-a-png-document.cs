using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class HorizontalRuleToPng
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape at the current cursor position.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Customize the appearance of the horizontal rule.
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // Centered on the page.
        format.WidthPercent = 70;                         // 70% of the page width.
        format.Height = 3;                                // Height of 3 points.
        format.Color = Color.Blue;                        // Blue color.
        format.NoShade = true;                            // Solid color, no 3‑D shading.

        // Prepare image save options for PNG output.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            HorizontalResolution = 96f, // Default horizontal DPI.
            VerticalResolution = 96f    // Default vertical DPI.
        };

        // Save the document as a PNG image.
        doc.Save("HorizontalRule.png", pngOptions);
    }
}
