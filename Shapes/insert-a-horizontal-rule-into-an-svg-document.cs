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

        // Insert a horizontal rule shape at the current cursor position.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Optional: customize the appearance of the horizontal rule.
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // Centered on the page.
        format.WidthPercent = 70;                         // 70% of the page width.
        format.Height = 3;                                // Height in points.
        format.Color = Color.Blue;                        // Solid blue color.
        format.NoShade = true;                            // Disable 3‑D shading.

        // Save the document as an SVG file.
        SvgSaveOptions svgOptions = new SvgSaveOptions
        {
            ShowPageBorder = false // Do not draw a border around the page.
        };
        doc.Save("HorizontalRule.svg", svgOptions);
    }
}
