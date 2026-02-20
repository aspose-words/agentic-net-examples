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

        // Optional: customize the appearance of the horizontal rule.
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // Centered on the page.
        format.WidthPercent = 80;                         // 80% of the page width.
        format.Height = 2;                                // 2 points thick.
        format.Color = Color.DarkGray;                    // Gray color.
        format.NoShade = true;                            // No 3‑D shading.

        // Save the document as an MHTML file.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
        doc.Save("HorizontalRule.mht", saveOptions);
    }
}
