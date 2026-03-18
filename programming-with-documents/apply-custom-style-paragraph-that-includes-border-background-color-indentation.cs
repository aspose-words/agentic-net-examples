using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a custom paragraph style.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");

        // ----- Indentation -----
        // Set left indent to 20 points.
        customStyle.ParagraphFormat.LeftIndent = 20;

        // ----- Border -----
        // Configure borders on all four sides.
        BorderCollection borders = customStyle.ParagraphFormat.Borders;
        borders.DistanceFromText = 5; // Space between text and border.

        borders[BorderType.Top].LineStyle = LineStyle.Single;
        borders[BorderType.Top].LineWidth = 2;
        borders[BorderType.Top].Color = Color.Blue;

        borders[BorderType.Bottom].LineStyle = LineStyle.Single;
        borders[BorderType.Bottom].LineWidth = 2;
        borders[BorderType.Bottom].Color = Color.Blue;

        borders[BorderType.Left].LineStyle = LineStyle.Single;
        borders[BorderType.Left].LineWidth = 2;
        borders[BorderType.Left].Color = Color.Blue;

        borders[BorderType.Right].LineStyle = LineStyle.Single;
        borders[BorderType.Right].LineWidth = 2;
        borders[BorderType.Right].Color = Color.Blue;

        // ----- Background shading -----
        // Apply a light yellow background to the paragraph.
        customStyle.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightYellow;

        // Apply the custom style to the current paragraph.
        builder.ParagraphFormat.Style = customStyle;
        builder.Writeln("This paragraph uses a custom style with border, shading, and indentation.");

        // Save the document to disk.
        doc.Save("CustomStyledParagraph.docx");
    }
}
