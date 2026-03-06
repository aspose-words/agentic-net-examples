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

        // Insert a JPEG image into the document (replace with your image path).
        string imagePath = "input.jpg";
        builder.InsertImage(imagePath);

        // Insert a horizontal rule shape after the image.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Optional: customize the appearance of the horizontal rule.
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // Centered on the page.
        format.WidthPercent = 80;                         // 80% of the page width.
        format.Height = 2;                                // Height in points.
        format.Color = Color.Blue;                        // Blue color.
        format.NoShade = true;                            // Solid color, no 3‑D shading.

        // Save the resulting document as a JPEG image.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg);
        doc.Save("output.jpg", saveOptions);
    }
}
