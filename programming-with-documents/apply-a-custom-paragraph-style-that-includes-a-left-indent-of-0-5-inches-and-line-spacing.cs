using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a custom paragraph style.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");

        // Set left indent to 0.5 inches (36 points).
        customStyle.ParagraphFormat.LeftIndent = 36.0;

        // Set line spacing to 1.5 lines.
        customStyle.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        // 1 line = 12 points, so 1.5 lines = 18 points.
        customStyle.ParagraphFormat.LineSpacing = 18.0;

        // Apply the custom style to a paragraph and write some text.
        builder.ParagraphFormat.StyleName = customStyle.Name;
        builder.Writeln("This paragraph uses a custom style with a left indent of 0.5 inches and line spacing of 1.5.");

        // Save the document.
        doc.Save("CustomParagraphStyle.docx");
    }
}
