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

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape at the current cursor position.
        Shape hrShape = builder.InsertHorizontalRule();

        // Access the HorizontalRuleFormat to customize appearance.
        HorizontalRuleFormat hrFormat = hrShape.HorizontalRuleFormat;

        // Center the rule, set its width to 70% of the page, height to 3 points,
        // use a blue color, and disable 3‑D shading.
        hrFormat.Alignment = HorizontalRuleAlignment.Center;
        hrFormat.WidthPercent = 70;
        hrFormat.Height = 3;
        hrFormat.Color = Color.Blue;
        hrFormat.NoShade = true;

        // Verify that the inserted shape is indeed a horizontal rule.
        if (!hrShape.IsHorizontalRule)
            throw new InvalidOperationException("The inserted shape is not a horizontal rule.");

        // Save the document to a file.
        doc.Save("HorizontalRule.docx");
    }
}
