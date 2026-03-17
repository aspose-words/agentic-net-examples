using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text before the rule.
        builder.Writeln("First section content.");

        // Insert a horizontal rule shape.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Access and customize the rule's formatting.
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center;   // Center the rule.
        format.WidthPercent = 80;                           // 80% of the page width.
        format.Height = 4;                                  // Thickness of 4 points.
        format.Color = Color.DarkBlue;                      // Blue color.
        format.NoShade = true;                              // Use solid color (no 3D shading).

        // Add text after the rule.
        builder.Writeln("Second section content.");

        // Save the document to a file.
        doc.Save("HorizontalRuleExample.docx");
    }
}
