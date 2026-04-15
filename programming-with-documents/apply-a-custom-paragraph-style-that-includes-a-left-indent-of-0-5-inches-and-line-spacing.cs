using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a custom paragraph style named "MyCustomStyle".
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");

        // Set left indent to 0.5 inches (36 points).
        customStyle.ParagraphFormat.LeftIndent = 36.0;

        // Set line spacing to 1.5 lines (multiple of default line spacing).
        customStyle.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        customStyle.ParagraphFormat.LineSpacing = 18.0; // 1.5 * 12 points

        // Insert a paragraph that uses the custom style.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ParagraphFormat.StyleName = "MyCustomStyle";
        builder.Writeln("This paragraph uses a custom style with a left indent of 0.5 inches and line spacing of 1.5.");

        // Save the document to the current directory.
        doc.Save("CustomStyle.docx");
    }
}
