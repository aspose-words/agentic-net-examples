using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a custom paragraph style.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");

        // Configure left border.
        Border leftBorder = customStyle.ParagraphFormat.Borders.Left;
        leftBorder.LineStyle = LineStyle.Single;
        leftBorder.LineWidth = 1.0;               // 1 point width
        leftBorder.Color = Color.LightGray;

        // Configure right border.
        Border rightBorder = customStyle.ParagraphFormat.Borders.Right;
        rightBorder.LineStyle = LineStyle.Single;
        rightBorder.LineWidth = 1.0;
        rightBorder.Color = Color.LightGray;

        // Set a light gray background shading.
        Shading shading = customStyle.ParagraphFormat.Shading;
        shading.Texture = TextureIndex.TextureSolid;
        shading.BackgroundPatternColor = Color.LightGray;

        // Apply the custom style to the current paragraph.
        builder.ParagraphFormat.Style = customStyle;
        builder.Writeln("This paragraph uses a custom style with left/right borders and a light gray background.");

        // Save the document.
        doc.Save("CustomParagraphStyle.docx");
    }
}
