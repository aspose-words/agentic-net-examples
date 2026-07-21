using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph of text.
        builder.Writeln("This is a quoted paragraph.");

        // Apply the built‑in Quote style to the current paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;

        // Increase the left indent for emphasis (e.g., 36 points = 0.5 inch).
        builder.ParagraphFormat.LeftIndent = 36;

        // Save the document to the local file system.
        const string outputFile = "QuoteStyle.docx";
        doc.Save(outputFile);
    }
}
