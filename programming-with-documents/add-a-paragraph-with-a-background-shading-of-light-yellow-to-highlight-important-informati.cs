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

        // Attach a DocumentBuilder to the document for easy content insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the paragraph shading to a light yellow background.
        // This will highlight the paragraph when the document is opened.
        builder.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightYellow;

        // Write the highlighted paragraph.
        builder.Writeln("This is an important paragraph highlighted with a light yellow background.");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HighlightedParagraph.docx");

        // Ensure the directory exists.
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

        // Save the document to the specified file.
        doc.Save(outputPath);
    }
}
