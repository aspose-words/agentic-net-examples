using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure left and right borders for the current paragraph.
        Border leftBorder = builder.ParagraphFormat.Borders.Left;
        leftBorder.LineStyle = LineStyle.Single;
        leftBorder.LineWidth = 1.0; // points
        leftBorder.Color = Color.Black;

        Border rightBorder = builder.ParagraphFormat.Borders.Right;
        rightBorder.LineStyle = LineStyle.Single;
        rightBorder.LineWidth = 1.0; // points
        rightBorder.Color = Color.Black;

        // Set a light gray background shading for the paragraph.
        Shading shading = builder.ParagraphFormat.Shading;
        shading.Texture = TextureIndex.TextureSolid;
        shading.BackgroundPatternColor = Color.LightGray;

        // Write a sample paragraph that will use the defined style.
        builder.Writeln("This paragraph has left/right borders and a light gray background.");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "CustomParagraphStyle.docx");
        doc.Save(outputPath);
    }
}
