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

        // Insert a horizontal rule shape
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Optional: customize the horizontal rule's appearance
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center;
        format.WidthPercent = 70;
        format.Height = 3;
        format.Color = Color.Blue;
        format.NoShade = true;

        // Save the document as an XPS file
        doc.Save("HorizontalRule.xps", SaveFormat.Xps);
    }
}
