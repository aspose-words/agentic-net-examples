using System;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write the paragraph that will be highlighted.
        builder.Writeln("Important information that needs to be highlighted.");

        // Apply light yellow background shading to the newly added paragraph.
        Paragraph paragraph = builder.CurrentParagraph;
        paragraph.ParagraphFormat.Shading.BackgroundPatternColor = Color.LightYellow;

        // Save the document to disk.
        string outputPath = "HighlightedParagraph.docx";
        doc.Save(outputPath);

        // Optional: confirm the file was created (no interactive output required).
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
