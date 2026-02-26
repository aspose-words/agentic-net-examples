using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class InsertHorizontalRuleToTxt
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some introductory text.
        builder.Writeln("Document with a horizontal rule:");

        // Insert the horizontal rule shape.
        Shape hrShape = builder.InsertHorizontalRule();

        // Optional: customize the appearance of the horizontal rule.
        HorizontalRuleFormat hrFormat = hrShape.HorizontalRuleFormat;
        hrFormat.Alignment = HorizontalRuleAlignment.Center;
        hrFormat.WidthPercent = 80;
        hrFormat.Height = 2;
        hrFormat.Color = System.Drawing.Color.DarkGray;
        hrFormat.NoShade = true;

        // Write text after the horizontal rule.
        builder.Writeln("Text after the horizontal rule.");

        // Prepare TXT save options (optional customizations can be set here).
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            // Ensure the document is saved as plain text.
            SaveFormat = SaveFormat.Text
        };

        // Save the document as a .txt file.
        doc.Save("HorizontalRuleDocument.txt", txtOptions);
    }
}
