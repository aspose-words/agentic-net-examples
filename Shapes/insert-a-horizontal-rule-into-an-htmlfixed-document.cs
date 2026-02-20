using System;
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

        // Insert a horizontal rule shape into the document.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Customize the appearance of the horizontal rule.
        HorizontalRuleFormat hrFormat = horizontalRule.HorizontalRuleFormat;
        hrFormat.Alignment = HorizontalRuleAlignment.Center;   // Center the rule.
        hrFormat.WidthPercent = 80;                           // 80% of the page width.
        hrFormat.Height = 2;                                  // Height in points.
        hrFormat.Color = Color.DarkGray;                      // Gray color.
        hrFormat.NoShade = true;                              // Solid color, no 3‑D shading.

        // Save the document as an HTML Fixed file.
        HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions
        {
            // Optional: set page alignment (default is Center).
            PageHorizontalAlignment = HtmlFixedPageHorizontalAlignment.Center
        };

        doc.Save("HorizontalRule.html", saveOptions);
    }
}
