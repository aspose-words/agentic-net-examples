using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document (or load an existing DOCM file)
        Document doc = new Document(); // new empty document
        // Document doc = new Document("input.docm"); // uncomment to load an existing file

        // Initialize a DocumentBuilder to work with the document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape at the current cursor position
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Optional: customize the appearance of the horizontal rule
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // center the rule
        format.WidthPercent = 70;                         // 70% of the page width
        format.Height = 3;                                // height in points
        format.Color = Color.Blue;                        // rule color
        format.NoShade = true;                            // solid color, no 3‑D shading

        // Save the document as a DOCM file
        doc.Save("output.docm", SaveFormat.Docm);
    }
}
