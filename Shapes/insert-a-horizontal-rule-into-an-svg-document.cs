using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape at the current cursor position.
        Shape hrShape = builder.InsertHorizontalRule();

        // Customize the horizontal rule formatting.
        HorizontalRuleFormat hrFormat = hrShape.HorizontalRuleFormat;
        hrFormat.Alignment = HorizontalRuleAlignment.Center; // Center the rule.
        hrFormat.WidthPercent = 70;                           // 70% of the page width.
        hrFormat.Height = 3;                                 // Height in points.
        hrFormat.Color = Color.Blue;                         // Blue color.
        hrFormat.NoShade = true;                             // Solid color, no 3D shading.

        // Verify that the shape is indeed a horizontal rule.
        if (!hrShape.IsHorizontalRule)
            throw new InvalidOperationException("Inserted shape is not a horizontal rule.");

        // Set up SVG save options (optional: you can adjust as needed).
        SvgSaveOptions svgOptions = new SvgSaveOptions
        {
            // Example: embed fonts as placed glyphs to ensure visual fidelity.
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs,
            // Example: remove page border for a cleaner SVG.
            ShowPageBorder = false,
            // Example: pretty format the SVG output.
            PrettyFormat = true
        };

        // Save the document as an SVG file.
        string outputPath = "HorizontalRule.svg";
        doc.Save(outputPath, svgOptions);
    }
}
