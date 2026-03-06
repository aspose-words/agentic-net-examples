using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape at the current cursor position.
        Shape hrShape = builder.InsertHorizontalRule();

        // Optional: customize the appearance of the horizontal rule.
        HorizontalRuleFormat hrFormat = hrShape.HorizontalRuleFormat;
        hrFormat.Alignment = HorizontalRuleAlignment.Center; // Centered on the page.
        hrFormat.WidthPercent = 70;                         // 70% of the page width.
        hrFormat.Height = 3;                                // Height of 3 points.
        hrFormat.Color = Color.Blue;                        // Blue color.
        hrFormat.NoShade = true;                            // Solid color, no 3‑D shading.

        // Save the document as a PDF file.
        doc.Save("HorizontalRule.pdf", SaveFormat.Pdf);
    }
}
