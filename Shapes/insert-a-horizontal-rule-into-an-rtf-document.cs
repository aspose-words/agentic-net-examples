using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape at the current cursor position.
        Shape hr = builder.InsertHorizontalRule();

        // Optional: customize the appearance of the horizontal rule.
        HorizontalRuleFormat format = hr.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // Centered on the page.
        format.WidthPercent = 70;                           // 70% of the page width.
        format.Height = 3;                                  // Height of 3 points.
        format.Color = Color.Blue;                         // Blue color.
        format.NoShade = true;                              // Solid color, no 3‑D shading.

        // Save the document as an RTF file.
        doc.Save("HorizontalRuleDocument.rtf");
    }
}
