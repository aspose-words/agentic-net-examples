using System;
using System.IO;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply light yellow background shading to the paragraph.
        builder.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightYellow;

        // Add the highlighted paragraph.
        builder.Writeln("Important: This information is highlighted with a light yellow background.");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "HighlightedParagraph.docx");
        doc.Save(outputPath);
    }
}
