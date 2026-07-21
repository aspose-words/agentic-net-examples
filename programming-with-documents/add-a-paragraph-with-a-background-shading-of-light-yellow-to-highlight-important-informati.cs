using System;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply light yellow background shading to the paragraph.
        builder.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightYellow;

        // Insert the highlighted paragraph.
        builder.Writeln("Important: Please review the highlighted information.");

        // Save the document.
        string outputFile = "HighlightedParagraph.docx";
        doc.Save(outputFile);
    }
}
