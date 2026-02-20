using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertHorizontalRuleExample
{
    static void Main()
    {
        // Load an existing HTML document.
        Document doc = new Document("input.html");

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule at the current cursor position.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Optional: customize the appearance of the horizontal rule.
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // Center the rule.
        format.WidthPercent = 80;                         // 80% of the page width.
        format.Height = 2;                                // Height in points.
        format.Color = Color.Gray;                        // Gray color.
        format.NoShade = true;                            // Solid color, no 3‑D shading.

        // Save the modified document back to HTML.
        doc.Save("output.html");
    }
}
