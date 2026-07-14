using System;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document for easy content insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a light yellow background shading to the current paragraph format.
        builder.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightYellow;

        // Insert the highlighted paragraph.
        builder.Writeln("Important: This information is highlighted with a light yellow background.");

        // (Optional) Clear shading for subsequent paragraphs.
        builder.ParagraphFormat.Shading.ClearFormatting();

        // Save the document to the local file system.
        string outputPath = "HighlightedParagraph.docx";
        doc.Save(outputPath);
    }
}
