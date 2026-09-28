using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a quoted paragraph.");

        // Get the paragraph that was just added.
        Paragraph paragraph = builder.CurrentParagraph;

        // Apply the built‑in "Quote" style.
        paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;

        // Increase left indent for emphasis (value is in points).
        paragraph.ParagraphFormat.LeftIndent = 20;

        // Save the document.
        string outputPath = "QuoteParagraph.docx";
        doc.Save(outputPath);
    }
}
