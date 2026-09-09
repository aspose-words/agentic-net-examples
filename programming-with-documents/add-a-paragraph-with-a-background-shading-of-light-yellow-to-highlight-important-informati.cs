using System;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder which will be used to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply light yellow background shading to the current paragraph.
        builder.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightYellow;

        // Write the highlighted paragraph.
        builder.Writeln("Important information highlighted with light yellow shading.");

        // Save the document to the local file system.
        string outputPath = "HighlightedParagraph.docx";
        doc.Save(outputPath);
    }
}
