using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class HorizontalRuleExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape at the current cursor position.
        Shape hr = builder.InsertHorizontalRule();

        // Access the HorizontalRuleFormat to customize appearance.
        HorizontalRuleFormat format = hr.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // Center the rule.
        format.WidthPercent = 70;                           // 70% of the page width.
        format.Height = 3;                                  // Height in points.
        format.Color = Color.Blue;                          // Rule color.
        format.NoShade = true;                              // Disable 3‑D shading.

        // Save the document to a .docx file.
        doc.Save("HorizontalRule.docx");
    }
}
