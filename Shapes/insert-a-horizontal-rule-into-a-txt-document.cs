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

        // Insert a horizontal rule shape.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Customize the horizontal rule's appearance (optional).
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center;
        format.WidthPercent = 70;
        format.Height = 3;
        format.Color = Color.Blue;
        format.NoShade = true;

        // Save the document as a plain‑text file.
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        doc.Save("HorizontalRule.txt", saveOptions);
    }
}
