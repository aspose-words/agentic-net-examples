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

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Optional: customize the appearance of the horizontal rule.
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center;
        format.WidthPercent = 80;   // 80% of the page width
        format.Height = 2;          // 2 points height
        format.Color = Color.Black;
        format.NoShade = true;      // Solid color without 3D shading

        // Prepare text save options (optional customization).
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            ParagraphBreak = "\r\n" // Use standard CRLF for paragraph breaks
        };

        // Save the document as a plain‑text file.
        doc.Save("HorizontalRule.txt", txtOptions);
    }
}
