using Aspose.Words;
using System;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for convenient editing.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in Quote style to the current paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;

        // Increase the left indent (e.g., 36 points = 0.5 inch) for emphasis.
        builder.ParagraphFormat.LeftIndent = 36;

        // Add the quoted text.
        builder.Writeln("“The only limit to our realization of tomorrow is our doubts of today.”");

        // Save the resulting document.
        doc.Save("QuoteParagraph.docx");
    }
}
