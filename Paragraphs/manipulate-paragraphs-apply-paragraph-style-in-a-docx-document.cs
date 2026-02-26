using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two paragraphs.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");

        // Apply a built‑in style to the first paragraph.
        Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
        firstParagraph.ParagraphFormat.StyleName = "Heading 1";

        // Apply a different built‑in style to the second paragraph.
        Paragraph secondParagraph = doc.FirstSection.Body.Paragraphs[1];
        secondParagraph.ParagraphFormat.StyleName = "Quote";

        // Save the document to a DOCX file.
        doc.Save("StyledParagraphs.docx");
    }
}
