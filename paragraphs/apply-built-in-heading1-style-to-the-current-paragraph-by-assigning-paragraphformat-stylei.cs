using System;
using Aspose.Words;

namespace ParagraphStyleExample
{
    class Program
    {
        static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph with some text.
            builder.Writeln("This is a heading paragraph.");

            // Apply built‑in Heading1 style to the current paragraph.
            Paragraph currentParagraph = builder.CurrentParagraph;
            currentParagraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            // Save the document.
            doc.Save("StyledParagraph.docx");
        }
    }
}
