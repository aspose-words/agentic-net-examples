using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in "Quote" style to the current paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;

        // Increase the left indent (e.g., 36 points = 0.5 inch) for emphasis.
        builder.ParagraphFormat.LeftIndent = 36;

        // Add some quoted text.
        builder.Writeln("The only limit to our realization of tomorrow is our doubts of today.");

        // Save the document to the local file system.
        doc.Save("QuoteStyle.docx");
    }
}
