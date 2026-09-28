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
        // Set font properties (optional).
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 12;
        // Configure paragraph formatting: border, background color, indentation.
        ParagraphFormat fmt = customStyle.ParagraphFormat;

        // Border settings.
        fmt.Borders.LineStyle = LineStyle.Single;
        fmt.Borders.Color = Color.DarkBlue;
        fmt.Borders.LineWidth = 2.0; // points

        // Background shading.
        fmt.Shading.BackgroundPatternColor = Color.LightYellow;

        // Indentation settings.
        fmt.LeftIndent = 20.0;          // points
        fmt.FirstLineIndent = 15.0;    // points

        // Apply the custom style to a new paragraph.
        builder.ParagraphFormat.Style = customStyle;
        builder.Writeln("This paragraph uses a custom style with a border, background color, and indentation.");

        // Save the document.
        string outputPath = "CustomStyleParagraph.docx";
        doc.Save(outputPath);
    }
}
