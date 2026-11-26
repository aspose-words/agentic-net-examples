using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with some text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph uses a locale‑independent style.");

        // Apply a built‑in style using the StyleIdentifier enumeration.
        // Heading1 is a standard style that is the same across locales.
        Paragraph paragraph = builder.CurrentParagraph;
        paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

        // Save the document to a file.
        doc.Save("Output.docx");

        // Indicate completion.
        Console.WriteLine("Document created with Heading1 style applied.");
    }
}
