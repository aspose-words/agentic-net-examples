using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a custom paragraph style to the document.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");

        // Configure the style's paragraph formatting.
        ParagraphFormat fmt = customStyle.ParagraphFormat;

        // Apply a solid border to all sides of the paragraph.
        foreach (Border border in fmt.Borders)
        {
            border.LineStyle = LineStyle.Single;
            border.LineWidth = 2.0;               // points
            border.Color = Color.Blue;
        }

        // Set a background color (shading) for the paragraph.
        fmt.Shading.BackgroundPatternColor = Color.LightYellow;

        // Define indentation values.
        fmt.LeftIndent = 20;        // points
        fmt.RightIndent = 20;       // points
        fmt.FirstLineIndent = 10;   // points

        // Insert a paragraph and apply the custom style.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ParagraphFormat.StyleName = "MyCustomStyle";
        builder.Writeln("This paragraph uses a custom style with a border, background color, and indentation.");

        // Save the document to a file.
        doc.Save("CustomStyleParagraph.docx");
    }
}
