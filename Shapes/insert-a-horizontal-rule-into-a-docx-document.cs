using System;
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

        // Insert a horizontal rule shape at the current cursor position.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Customize the appearance of the horizontal rule (optional).
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // Centered on the page.
        format.WidthPercent = 80;                         // 80% of the page width.
        format.Height = 2;                                // Height in points.
        format.Color = Color.DarkGray;                    // Rule color.
        format.NoShade = true;                            // Solid color, no 3‑D shading.

        // Save the document to a DOCX file.
        doc.Save("HorizontalRule.docx");
    }
}
