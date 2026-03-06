using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape at the current cursor position
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Optional: customize the appearance of the horizontal rule
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // Centered alignment
        format.WidthPercent = 70;                         // 70% of the page width
        format.Height = 3;                                // Height in points
        format.Color = Color.Blue;                        // Blue color
        format.NoShade = true;                            // Solid color without 3D shading

        // Save the document to a file
        doc.Save("HorizontalRule.docx");
    }
}
