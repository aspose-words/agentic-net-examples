using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document (DOTX template)
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape into the document
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Optional: customize the horizontal rule's appearance
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // Center the rule
        format.WidthPercent = 70;                         // 70% of the page width
        format.Height = 3;                                // Height in points
        format.Color = Color.Blue;                        // Blue color
        format.NoShade = true;                            // Solid color, no 3D shading

        // Save the document as a DOTX template
        doc.Save("HorizontalRuleTemplate.dotx", SaveFormat.Dotx);
    }
}
