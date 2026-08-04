using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in "Quote" style to the current paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;

        // Increase the left indent (in points) for visual emphasis.
        builder.ParagraphFormat.LeftIndent = 30; // 30 points

        // Add some quoted text.
        builder.Writeln("“The only limit to our realization of tomorrow is our doubts of today.” – Franklin D. Roosevelt");

        // Save the document.
        const string outputFile = "QuoteStyleExample.docx";
        doc.Save(outputFile);
    }
}
