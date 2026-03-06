using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using System.Drawing;

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
        format.Alignment = HorizontalRuleAlignment.Center; // Centered on the page.
        format.WidthPercent = 80;                         // 80% of the page width.
        format.Height = 2;                                // 2 points thick.
        format.Color = Color.DarkGray;                    // Gray color.
        format.NoShade = true;                            // Solid color, no 3‑D shading.

        // Save the document as an SVG file.
        SvgSaveOptions svgOptions = new SvgSaveOptions
        {
            ShowPageBorder = false,                       // No outer page border.
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs // Preserve selectable text.
        };
        doc.Save("HorizontalRule.svg", svgOptions);
    }
}
