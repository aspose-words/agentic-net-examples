using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Customize the horizontal rule formatting.
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // Center the rule.
        format.WidthPercent = 70;                           // 70% of the page width.
        format.Height = 3;                                  // Height in points.
        format.Color = Color.Blue;                          // Blue color.
        format.NoShade = true;                              // Disable 3‑D shading.

        // Save the document in WORDML (XML) format.
        doc.Save("HorizontalRule.xml", SaveFormat.WordML);
    }
}
