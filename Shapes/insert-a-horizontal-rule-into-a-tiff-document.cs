using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Customize the horizontal rule's appearance (optional).
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center;
        format.WidthPercent = 80;   // 80% of the page width.
        format.Height = 2;          // Height in points.
        format.Color = Color.Black;
        format.NoShade = true;      // Solid color without 3D shading.

        // Save the document as a TIFF image.
        doc.Save("HorizontalRule.tiff", SaveFormat.Tiff);
    }
}
