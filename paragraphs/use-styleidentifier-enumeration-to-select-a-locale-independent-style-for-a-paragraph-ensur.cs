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
        builder.Writeln("This paragraph uses a locale‑independent built‑in style.");

        // Retrieve the paragraph that was just added.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Apply a built‑in style using the StyleIdentifier enumeration.
        // Heading1 is a locale‑independent style that exists in all Word documents.
        paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

        // Save the document to the output file.
        doc.Save("Output.docx");
    }
}
