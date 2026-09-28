using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with Arabic text.
        builder.Writeln("مرحبا بالعالم");

        // Enable right‑to‑left direction for the paragraph.
        Paragraph para = builder.CurrentParagraph;
        para.ParagraphFormat.Bidi = true;

        // Save the document.
        doc.Save("Output.docx");
    }
}
