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

        // Create a DocumentBuilder to facilitate editing.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a custom paragraph style.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");

        // Configure left border.
        Border leftBorder = customStyle.ParagraphFormat.Borders.Left;
        leftBorder.LineStyle = LineStyle.Single;
        leftBorder.LineWidth = 1.0;
        leftBorder.Color = Color.Black;

        // Configure right border.
        Border rightBorder = customStyle.ParagraphFormat.Borders.Right;
        rightBorder.LineStyle = LineStyle.Single;
        rightBorder.LineWidth = 1.0;
        rightBorder.Color = Color.Black;

        // Set a light gray background shading.
        Shading shading = customStyle.ParagraphFormat.Shading;
        shading.BackgroundPatternColor = Color.LightGray;

        // Apply the custom style to the next paragraph.
        builder.ParagraphFormat.StyleName = customStyle.Name;
        builder.Writeln("This paragraph uses a custom style with left/right borders and a light gray background.");

        // Save the document.
        doc.Save("CustomParagraphStyle.docx");
    }
}
