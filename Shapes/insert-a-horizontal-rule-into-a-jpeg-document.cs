using System;
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
        format.Alignment = HorizontalRuleAlignment.Center;   // Centered on the page.
        format.WidthPercent = 80;                           // 80% of the page width.
        format.Height = 2;                                  // Height in points.
        format.Color = Color.Black;                         // Black color.
        format.NoShade = true;                              // Solid color, no 3‑D shading.

        // Save the document as a JPEG image.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            HorizontalResolution = 96f,
            VerticalResolution = 96f
        };

        doc.Save("HorizontalRule.jpg", saveOptions);
    }
}
