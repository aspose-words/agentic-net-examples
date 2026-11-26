using System;
using System.Drawing;
using System.IO;
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

        // Set left and right borders.
        Border leftBorder = customStyle.ParagraphFormat.Borders[BorderType.Left];
        leftBorder.LineStyle = LineStyle.Single;
        leftBorder.Color = Color.Black;
        leftBorder.LineWidth = 1.0; // points

        Border rightBorder = customStyle.ParagraphFormat.Borders[BorderType.Right];
        rightBorder.LineStyle = LineStyle.Single;
        rightBorder.Color = Color.Black;
        rightBorder.LineWidth = 1.0; // points

        // Set a light gray background shading.
        customStyle.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightGray;

        // Apply the custom style to a paragraph.
        builder.ParagraphFormat.Style = customStyle;
        builder.Writeln("This paragraph uses a custom style with left/right borders and a light gray background.");

        // Save the document.
        string outputPath = "CustomStyle.docx";
        doc.Save(outputPath);

        // Verify that the file was created (optional, no output required).
        if (File.Exists(outputPath))
        {
            // Document created successfully.
        }
    }
}
