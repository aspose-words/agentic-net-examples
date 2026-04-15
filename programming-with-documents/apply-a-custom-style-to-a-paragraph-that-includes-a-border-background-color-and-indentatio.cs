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

        // Create a DocumentBuilder which will be used to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Define a custom paragraph style named "CustomStyle".
        // -----------------------------------------------------------------
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "CustomStyle");

        // Set a border on all four sides of the paragraph.
        BorderCollection borders = customStyle.ParagraphFormat.Borders;
        foreach (Border border in new[] { borders.Top, borders.Bottom, borders.Left, borders.Right })
        {
            border.LineStyle = LineStyle.Single;
            border.LineWidth = 1.5;               // points
            border.Color = Color.DarkBlue;
        }

        // Set a background shading (fill) color.
        customStyle.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightYellow;

        // Set indentation values.
        customStyle.ParagraphFormat.LeftIndent = 20;          // points
        customStyle.ParagraphFormat.RightIndent = 20;         // points
        customStyle.ParagraphFormat.FirstLineIndent = 10;     // points

        // -----------------------------------------------------------------
        // Apply the custom style to a new paragraph and write some text.
        // -----------------------------------------------------------------
        builder.ParagraphFormat.Style = customStyle;
        builder.Writeln("This paragraph uses a custom style with borders, shading, and indentation.");

        // Save the document to the local file system.
        doc.Save("CustomStyleParagraph.docx");
    }
}
