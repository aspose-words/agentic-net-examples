using System.Text;
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

        // Customize the appearance of the horizontal rule.
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // Centered.
        format.WidthPercent = 70;                         // 70% of page width.
        format.Height = 3;                                // Height in points.
        format.Color = Color.Blue;                        // Blue color.
        format.NoShade = true;                            // Solid color, no 3‑D shading.

        // Prepare EPUB save options.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            SaveFormat = SaveFormat.Epub,   // Export as EPUB.
            Encoding = Encoding.UTF8,       // Use UTF‑8 encoding.
            DocumentSplitCriteria = DocumentSplitCriteria.None // Single HTML part.
        };

        // Save the document as an EPUB file.
        doc.Save("HorizontalRule.epub", saveOptions);
    }
}
