using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing HTML document.
        Document doc = new Document("input.html");

        // Create a DocumentBuilder for the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape at the current cursor position.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Optional: customize the appearance of the horizontal rule.
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // Center the rule.
        format.WidthPercent = 80;                         // 80% of the page width.
        format.Height = 2;                                // Height in points.
        format.Color = Color.Black;                       // Solid black line.
        format.NoShade = true;                            // Disable 3‑D shading.

        // Save the modified document back to HTML format.
        doc.Save("output.html");
    }
}
