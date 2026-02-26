using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class InsertHorizontalRuleExample
{
    static void Main()
    {
        // Load an existing HTML document.
        Document doc = new Document("input.html");

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document (or any desired position).
        builder.MoveToDocumentEnd();

        // Insert a horizontal rule shape.
        Shape hrShape = builder.InsertHorizontalRule();

        // Optional: customize the horizontal rule's appearance.
        HorizontalRuleFormat hrFormat = hrShape.HorizontalRuleFormat;
        hrFormat.Alignment = HorizontalRuleAlignment.Center; // Centered on the page.
        hrFormat.WidthPercent = 80;                         // 80% of the page width.
        hrFormat.Height = 2;                                // Height in points.
        hrFormat.Color = Color.DarkGray;                    // Rule color.
        hrFormat.NoShade = true;                            // Solid color, no 3‑D shading.

        // Save the modified document back to HTML format.
        doc.Save("output.html");
    }
}
