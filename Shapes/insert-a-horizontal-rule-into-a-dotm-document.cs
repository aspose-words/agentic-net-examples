using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

namespace AsposeWordsHorizontalRuleExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new empty document.
            Document doc = new Document();

            // Initialize DocumentBuilder to work with the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a horizontal rule shape at the current cursor position.
            Shape horizontalRule = builder.InsertHorizontalRule();

            // Optional: customize the appearance of the horizontal rule.
            HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
            format.Alignment = HorizontalRuleAlignment.Center;   // Center the rule.
            format.WidthPercent = 70;                           // 70% of the page width.
            format.Height = 3;                                  // Height in points.
            format.Color = Color.Blue;                          // Blue color.
            format.NoShade = true;                              // Solid color, no 3D shading.

            // Save the document as a macro-enabled template (.dotm).
            doc.Save("HorizontalRuleTemplate.dotm");
        }
    }
}
