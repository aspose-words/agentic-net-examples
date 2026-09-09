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

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure paragraph formatting.

        // 1. Add a solid border around the paragraph.
        // The Borders property returns a BorderCollection; set its properties to affect all sides.
        BorderCollection borders = builder.ParagraphFormat.Borders;
        borders.LineStyle = LineStyle.Single;
        borders.Color = Color.DarkBlue;
        borders.LineWidth = 2.0; // points

        // 2. Set a background shading color.
        builder.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightYellow;

        // 3. Apply left and right indentation (points).
        builder.ParagraphFormat.LeftIndent = 30;
        builder.ParagraphFormat.RightIndent = 30;

        // Write the paragraph text.
        builder.Writeln("This paragraph has a custom style with a border, background color, and indentation.");

        // Save the document.
        doc.Save("CustomStyledParagraph.docx");
    }
}
