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

        // Add a custom paragraph style.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");

        // Set border for the style.
        customStyle.ParagraphFormat.Borders.LineStyle = LineStyle.Single;
        customStyle.ParagraphFormat.Borders.Color = Color.Black;
        customStyle.ParagraphFormat.Borders.LineWidth = 1.0; // points

        // Set background color for the style.
        customStyle.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightYellow;

        // Set indentation for the style.
        customStyle.ParagraphFormat.LeftIndent = 20;   // points
        customStyle.ParagraphFormat.RightIndent = 20;  // points
        customStyle.ParagraphFormat.FirstLineIndent = 10; // points

        // Use DocumentBuilder to add a paragraph with the custom style.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ParagraphFormat.Style = customStyle;
        builder.Writeln("This paragraph uses a custom style with a border, background color, and indentation.");

        // Save the document to disk.
        string outputPath = "CustomStyle.docx";
        doc.Save(outputPath);
    }
}
