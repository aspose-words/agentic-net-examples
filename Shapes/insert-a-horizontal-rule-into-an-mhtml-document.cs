using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertHorizontalRuleMhtml
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape at the current cursor position.
        Shape hrShape = builder.InsertHorizontalRule();

        // Optional: customize the appearance of the horizontal rule.
        HorizontalRuleFormat hrFormat = hrShape.HorizontalRuleFormat;
        hrFormat.Alignment = HorizontalRuleAlignment.Center; // center the rule
        hrFormat.WidthPercent = 70;                         // 70% of the page width
        hrFormat.Height = 3;                                // 3 points height
        hrFormat.Color = Color.Blue;                        // blue color
        hrFormat.NoShade = true;                            // solid color, no 3‑D shading

        // Save the document as MHTML.
        doc.Save("HorizontalRule.mhtml", SaveFormat.Mhtml);
    }
}
