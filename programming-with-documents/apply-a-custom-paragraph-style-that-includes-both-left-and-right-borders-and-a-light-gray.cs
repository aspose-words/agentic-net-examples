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
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");

        // Configure left border.
        Border leftBorder = customStyle.ParagraphFormat.Borders.Left;
        leftBorder.LineStyle = LineStyle.Single;
        leftBorder.LineWidth = 1.0; // points
        leftBorder.Color = Color.Black;

        // Configure right border.
        Border rightBorder = customStyle.ParagraphFormat.Borders.Right;
        rightBorder.LineStyle = LineStyle.Single;
        rightBorder.LineWidth = 1.0; // points
        rightBorder.Color = Color.Black;

        // Set a light gray background shading.
        customStyle.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightGray;

        // Apply the custom style to the current paragraph.
        builder.ParagraphFormat.Style = customStyle;

        // Write sample text that will use the style.
        builder.Writeln("This paragraph is formatted with left/right borders and a light gray background.");

        // Save the document to the local file system.
        doc.Save("CustomParagraphStyle.docx");
    }
}
