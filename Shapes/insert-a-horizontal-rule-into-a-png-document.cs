using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape at the current cursor position
        Shape hrShape = builder.InsertHorizontalRule();

        // Optional: customize the horizontal rule appearance
        HorizontalRuleFormat hrFormat = hrShape.HorizontalRuleFormat;
        hrFormat.Alignment = HorizontalRuleAlignment.Center;   // center the rule
        hrFormat.WidthPercent = 80;                           // 80% of page width
        hrFormat.Height = 2;                                  // 2 points thick
        hrFormat.Color = Color.DarkGray;                      // gray color
        hrFormat.NoShade = true;                              // solid color, no 3D shading

        // Save the document as a PNG image (each page will be rendered to a separate PNG file)
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png);
        doc.Save("HorizontalRule.png", saveOptions);
    }
}
