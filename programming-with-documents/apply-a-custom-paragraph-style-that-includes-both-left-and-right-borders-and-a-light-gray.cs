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

        // Define a custom paragraph style.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "CustomBorderGrayBackground");

        // Set left border.
        customStyle.ParagraphFormat.Borders[BorderType.Left].LineStyle = LineStyle.Single;
        customStyle.ParagraphFormat.Borders[BorderType.Left].Color = Color.Black;
        customStyle.ParagraphFormat.Borders[BorderType.Left].LineWidth = 1.0;

        // Set right border.
        customStyle.ParagraphFormat.Borders[BorderType.Right].LineStyle = LineStyle.Single;
        customStyle.ParagraphFormat.Borders[BorderType.Right].Color = Color.Black;
        customStyle.ParagraphFormat.Borders[BorderType.Right].LineWidth = 1.0;

        // Set light gray background shading.
        customStyle.ParagraphFormat.Shading.Texture = TextureIndex.TextureSolid;
        customStyle.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightGray;

        // Write a paragraph and apply the custom style.
        builder.Writeln("This paragraph uses a custom style with left/right borders and a light gray background.");
        builder.CurrentParagraph.ParagraphFormat.Style = customStyle;

        // Save the document.
        string outputPath = "CustomStyle.docx";
        doc.Save(outputPath);
    }
}
