using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class InsertHorizontalRuleExample
{
    static void Main()
    {
        // Create a new empty Word document.
        Document doc = new Document();

        // Initialize DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Optional: add some introductory text.
        builder.Writeln("Below is a horizontal rule:");

        // Insert the horizontal rule shape.
        Shape hrShape = builder.InsertHorizontalRule();

        // Customize the appearance of the horizontal rule (optional).
        HorizontalRuleFormat hrFormat = hrShape.HorizontalRuleFormat;
        hrFormat.Alignment = HorizontalRuleAlignment.Center; // Centered on the page.
        hrFormat.WidthPercent = 80;                         // 80% of the page width.
        hrFormat.Height = 2;                                // Thickness of the rule.
        hrFormat.Color = Color.DarkGray;                    // Color of the rule.
        hrFormat.NoShade = true;                            // Solid color, no 3D shading.

        // Optional: add text after the rule.
        builder.Writeln("Text after the horizontal rule.");

        // Save the document as a Markdown file.
        doc.Save("HorizontalRuleExample.md", SaveFormat.Markdown);
    }
}
