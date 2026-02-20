using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Customize the horizontal rule's formatting.
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // Center the rule.
        format.WidthPercent = 80;                           // 80% of the page width.
        format.Height = 2;                                  // Height in points.
        format.Color = Color.Black;                         // Rule color.
        format.NoShade = true;                              // Disable 3‑D shading.

        // Save the document as a macro‑enabled DOCM file.
        doc.Save("HorizontalRule.docm", SaveFormat.Docm);
    }
}
