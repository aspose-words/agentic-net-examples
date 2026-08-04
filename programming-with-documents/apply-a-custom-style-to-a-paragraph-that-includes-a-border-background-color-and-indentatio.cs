using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure paragraph formatting: left indent, border, and background shading.
        ParagraphFormat paraFormat = builder.ParagraphFormat;

        // Indentation (20 points from the left margin).
        paraFormat.LeftIndent = 20;

        // Apply a single line border to all sides.
        paraFormat.Borders.LineStyle = LineStyle.Single;
        paraFormat.Borders.Color = Color.DarkBlue;
        paraFormat.Borders.LineWidth = 2.0;

        // Set background color (light yellow).
        paraFormat.Shading.BackgroundPatternColor = Color.LightYellow;

        // Write the paragraph text.
        builder.Writeln("This paragraph has a custom style with a border, background color, and indentation.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "CustomStyledParagraph.docx");
        doc.Save(outputPath);
    }
}
