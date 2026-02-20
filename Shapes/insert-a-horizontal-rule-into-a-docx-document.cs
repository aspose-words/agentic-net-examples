using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertHorizontalRuleExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape at the current cursor position.
        Shape hrShape = builder.InsertHorizontalRule();

        // Optional: customize the appearance of the horizontal rule.
        HorizontalRuleFormat hrFormat = hrShape.HorizontalRuleFormat;
        hrFormat.Alignment = HorizontalRuleAlignment.Center; // Center the rule.
        hrFormat.WidthPercent = 70;                           // 70% of the page width.
        hrFormat.Height = 3;                                 // Height in points.
        hrFormat.Color = Color.Blue;                         // Rule color.
        hrFormat.NoShade = true;                             // Use solid color without 3‑D shading.

        // Save the document to a DOCX file.
        doc.Save("HorizontalRule.docx");
    }
}
