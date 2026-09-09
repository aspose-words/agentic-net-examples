using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a custom paragraph style.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");

        // Set left indent to 0.5 inches (36 points).
        customStyle.ParagraphFormat.LeftIndent = 36.0;

        // Set line spacing to 1.5 lines.
        customStyle.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        customStyle.ParagraphFormat.LineSpacing = 18.0; // 12 points * 1.5 = 18 points

        // Use DocumentBuilder to write a paragraph with the custom style.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ParagraphFormat.StyleName = customStyle.Name;
        builder.Writeln("This paragraph uses a custom style with a left indent of 0.5 inches and line spacing of 1.5.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "CustomParagraphStyle.docx");
        doc.Save(outputPath);
    }
}
