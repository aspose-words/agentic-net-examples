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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a custom paragraph style.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");

        // Set border for the style (single line, black, 1 point).
        foreach (Border border in customStyle.ParagraphFormat.Borders)
        {
            border.LineStyle = LineStyle.Single;
            border.Color = Color.Black;
            border.LineWidth = 1.0;
        }

        // Set background color for the style.
        customStyle.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightYellow;

        // Set indentation (left indent 20 points, first line indent 10 points).
        customStyle.ParagraphFormat.LeftIndent = 20.0;
        customStyle.ParagraphFormat.FirstLineIndent = 10.0;

        // Apply the custom style to a new paragraph.
        builder.ParagraphFormat.Style = customStyle;
        builder.Writeln("This paragraph uses a custom style with border, background color, and indentation.");

        // Save the document.
        string outputPath = "CustomStyle.docx";
        doc.Save(outputPath);
    }
}
