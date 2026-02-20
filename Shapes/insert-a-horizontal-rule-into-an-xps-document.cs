using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape at the current cursor position.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Customize the formatting of the horizontal rule.
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // Center alignment.
        format.WidthPercent = 70;                           // 70% of the page width.
        format.Height = 3;                                 // Height of 3 points.
        format.Color = Color.Blue;                         // Blue color.
        format.NoShade = true;                             // No 3‑D shading.

        // Save the document as an XPS file using default XpsSaveOptions.
        XpsSaveOptions saveOptions = new XpsSaveOptions();
        doc.Save("HorizontalRule.xps", saveOptions);
    }
}
