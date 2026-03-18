using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a custom paragraph style.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");

        // Set left indent to 0.5 inches (0.5 * 72 points per inch = 36 points).
        customStyle.ParagraphFormat.LeftIndent = 36.0;

        // Set line spacing to 1.5 (multiple of default line spacing).
        customStyle.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        customStyle.ParagraphFormat.LineSpacing = 1.5;

        // Use DocumentBuilder to write a paragraph that uses the custom style.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ParagraphFormat.Style = customStyle;
        builder.Writeln("This paragraph uses a custom style with a left indent of 0.5 inches and line spacing of 1.5.");

        // Save the document.
        string outputPath = "CustomParagraphStyle.docx";
        doc.Save(outputPath);
    }
}
