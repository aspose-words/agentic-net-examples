using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a custom paragraph style.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        // Set left indent to 0.5 inches (convert inches to points).
        customStyle.ParagraphFormat.LeftIndent = ConvertUtil.InchToPoint(0.5);
        // Set line spacing to 1.5 (multiple line spacing).
        customStyle.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        customStyle.ParagraphFormat.LineSpacing = 1.5;

        // Apply the custom style to a paragraph.
        builder.ParagraphFormat.StyleName = "MyCustomStyle";
        builder.Writeln("This paragraph uses a custom style with a left indent of 0.5 inches and 1.5 line spacing.");

        // Save the document.
        string outputPath = "CustomStyle.docx";
        doc.Save(outputPath);
    }
}
