using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        // Ensure the document contains at least one section and one paragraph.
        doc.EnsureMinimum();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write the first paragraph (default formatting).
        builder.Writeln("This is the original paragraph.");

        // Insert a new empty paragraph after the current one.
        Paragraph formattedParagraph = builder.InsertParagraph();

        // Apply paragraph formatting:
        // - Center the text.
        // - Set a first‑line indent of 20 points.
        // - Add 10 points of space after the paragraph.
        formattedParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        formattedParagraph.ParagraphFormat.FirstLineIndent = 20;
        formattedParagraph.ParagraphFormat.SpaceAfter = 10;

        // Add text to the newly formatted paragraph.
        builder.Writeln("This paragraph is centered with indentation.");

        // Save the document to a DOCX file.
        doc.Save("FormattedParagraph.docx");
    }
}
