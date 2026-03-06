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

        // Initialize a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a horizontal rule shape at the current cursor position.
        Shape horizontalRule = builder.InsertHorizontalRule();

        // Customize the appearance of the horizontal rule (optional).
        HorizontalRuleFormat format = horizontalRule.HorizontalRuleFormat;
        format.Alignment = HorizontalRuleAlignment.Center; // Center the rule.
        format.WidthPercent = 80;                         // 80% of the page width.
        format.Height = 2;                                // Height of 2 points.
        format.Color = Color.DarkGray;                    // Gray color.
        format.NoShade = true;                            // Solid color, no 3‑D shading.

        // Add some text after the horizontal rule.
        builder.Writeln("Content after the horizontal rule.");

        // Configure save options for EPUB output.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.SaveFormat = SaveFormat.Epub;               // Specify EPUB format.
        saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.None; // Keep content in a single part.

        // Save the document as an EPUB file.
        doc.Save("HorizontalRule.epub", saveOptions);
    }
}
