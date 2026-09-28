using System;
using System.Drawing;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set paragraph shading to light yellow.
        builder.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightYellow;

        // Add the highlighted paragraph.
        builder.Writeln("Important information highlighted with light yellow background.");

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HighlightedParagraph.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            // Optionally reopen to ensure it loads without error.
            Document loadedDoc = new Document(outputPath);
        }
    }
}
