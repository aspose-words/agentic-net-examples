using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Customize the horizontal rule's appearance (optional).
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // Centered.
        format.WidthPercent = 70;                           // 70% of page width.
        format.Height = 3;                                  // 3 points tall.
        format.Color = Color.Blue;                         // Blue color.
        format.NoShade = true;                              // Solid color, no 3‑D shading.

        // Save the document as a Markdown file.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
        doc.Save("HorizontalRule.md", saveOptions);
    }
}
