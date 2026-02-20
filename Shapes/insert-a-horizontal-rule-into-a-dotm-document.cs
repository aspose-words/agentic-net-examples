using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing DOTM template.
        Document doc = new Document(@"Template.dotm");

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule at the current cursor position.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Optional: customize the appearance of the horizontal rule.
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // Center the rule.
        format.WidthPercent = 80;                           // 80% of the page width.
        format.Height = 2;                                 // Height in points.
        format.Color = Color.DarkGray;                     // Rule color.
        format.NoShade = true;                             // Disable 3D shading.

        // Save the modified document back as a DOTM file.
        doc.Save(@"Result.dotm", SaveFormat.Dotm);
    }
}
