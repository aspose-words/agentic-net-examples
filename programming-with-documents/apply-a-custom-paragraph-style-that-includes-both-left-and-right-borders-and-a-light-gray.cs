using System;
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
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "CustomBorderGray");
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 12;

        // Configure left border.
        customStyle.ParagraphFormat.Borders.Left.LineStyle = LineStyle.Single;
        customStyle.ParagraphFormat.Borders.Left.LineWidth = 1.0;
        customStyle.ParagraphFormat.Borders.Left.Color = System.Drawing.Color.Black;

        // Configure right border.
        customStyle.ParagraphFormat.Borders.Right.LineStyle = LineStyle.Single;
        customStyle.ParagraphFormat.Borders.Right.LineWidth = 1.0;
        customStyle.ParagraphFormat.Borders.Right.Color = System.Drawing.Color.Black;

        // Set a light gray background shading.
        customStyle.ParagraphFormat.Shading.BackgroundPatternColor = System.Drawing.Color.LightGray;

        // Apply the custom style to a new paragraph.
        builder.ParagraphFormat.StyleName = customStyle.Name;
        builder.Writeln("This paragraph uses a custom style with left/right borders and a light gray background.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomParagraphStyle.docx");
        doc.Save(outputPath);
    }
}
