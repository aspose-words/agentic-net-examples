using System;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document for easy content insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the paragraph shading to a light yellow color.
        // This will highlight the paragraph background.
        builder.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightYellow;

        // Write the highlighted paragraph.
        builder.Writeln("This is an important paragraph highlighted with a light yellow background.");

        // Save the document to the local file system.
        doc.Save("HighlightedParagraph.docx");
    }
}
