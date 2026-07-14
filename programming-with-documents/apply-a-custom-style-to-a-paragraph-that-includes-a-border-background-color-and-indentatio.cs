using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a custom paragraph style named "MyCustomStyle".
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");

        // Configure the style's paragraph formatting.
        // 1. Border – apply a single blue line to all sides.
        BorderCollection borders = customStyle.ParagraphFormat.Borders;
        borders.LineStyle = LineStyle.Single;
        borders.Color = Color.Blue;
        borders.LineWidth = 2.0;

        // 2. Background shading – light yellow fill.
        customStyle.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightYellow;

        // 3. Indentation – 20 points on left and right.
        customStyle.ParagraphFormat.LeftIndent = 20;
        customStyle.ParagraphFormat.RightIndent = 20;

        // Use DocumentBuilder to insert a paragraph that uses the custom style.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ParagraphFormat.StyleName = "MyCustomStyle";
        builder.Writeln("This paragraph demonstrates a custom style with a border, background color, and indentation.");

        // Save the document to a file in the same folder as the executable.
        string outputPath = "CustomStyleParagraph.docx";
        doc.Save(outputPath);
    }
}
