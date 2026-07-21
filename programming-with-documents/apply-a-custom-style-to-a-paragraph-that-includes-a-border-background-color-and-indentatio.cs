using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // 1. Create a new blank document.
        Document doc = new Document();

        // 2. Add a custom paragraph style to the document.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");

        // 3. Configure the style's paragraph formatting.

        // 3.1. Border: single blue line, 2 points wide, applied to all sides.
        BorderCollection borders = customStyle.ParagraphFormat.Borders;
        borders[BorderType.Left].LineStyle = LineStyle.Single;
        borders[BorderType.Left].Color = Color.Blue;
        borders[BorderType.Left].LineWidth = 2.0;

        borders[BorderType.Right].LineStyle = LineStyle.Single;
        borders[BorderType.Right].Color = Color.Blue;
        borders[BorderType.Right].LineWidth = 2.0;

        borders[BorderType.Top].LineStyle = LineStyle.Single;
        borders[BorderType.Top].Color = Color.Blue;
        borders[BorderType.Top].LineWidth = 2.0;

        borders[BorderType.Bottom].LineStyle = LineStyle.Single;
        borders[BorderType.Bottom].Color = Color.Blue;
        borders[BorderType.Bottom].LineWidth = 2.0;

        // 3.2. Background shading: light yellow.
        customStyle.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightYellow;

        // 3.3. Indentation: left/right 20 points, first line 10 points.
        customStyle.ParagraphFormat.LeftIndent = 20;
        customStyle.ParagraphFormat.RightIndent = 20;
        customStyle.ParagraphFormat.FirstLineIndent = 10;

        // 4. Insert a paragraph that uses the custom style.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ParagraphFormat.StyleName = "MyCustomStyle";
        builder.Writeln("This paragraph uses a custom style with a border, background color, and indentation.");

        // 5. Save the document.
        doc.Save("CustomStyleParagraph.docx");
    }
}
