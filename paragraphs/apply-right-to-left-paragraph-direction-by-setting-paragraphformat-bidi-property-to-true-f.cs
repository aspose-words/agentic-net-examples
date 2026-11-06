using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to insert a paragraph with Arabic text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("مرحبا بالعالم"); // Arabic: "Hello World"

        // The paragraph that was just created is the current paragraph of the builder.
        Paragraph arabicParagraph = builder.CurrentParagraph;

        // Apply right‑to‑left direction for the paragraph.
        arabicParagraph.ParagraphFormat.Bidi = true;

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
