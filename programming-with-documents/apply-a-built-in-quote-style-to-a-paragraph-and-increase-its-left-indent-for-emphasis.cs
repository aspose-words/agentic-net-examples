using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph of text.
        builder.Writeln("This is a quoted paragraph.");

        // Apply the built‑in Quote style to the current paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;

        // Increase the left indent (e.g., 36 points = 0.5 inch) for emphasis.
        builder.ParagraphFormat.LeftIndent = 36;

        // Save the document to the local file system.
        doc.Save("QuoteStyle.docx");
    }
}
