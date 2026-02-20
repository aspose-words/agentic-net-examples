using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule at the current cursor position.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Customize the appearance of the horizontal rule (optional).
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // Centered on the page.
        format.WidthPercent = 80;                           // 80% of the page width.
        format.Height = 2;                                  // 2 points tall.
        format.Color = Color.Black;                         // Black color.
        format.NoShade = true;                              // No 3‑D shading.

        // Save the document in RTF format.
        RtfSaveOptions rtfOptions = new RtfSaveOptions();
        doc.Save("HorizontalRule.rtf", rtfOptions);
    }
}
