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

        // Title of the FAQ.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
        builder.Writeln("Frequently Asked Questions");
        builder.Writeln(); // Blank line after title.

        // Define the FAQ entries.
        var faqs = new (string Question, string Answer)[]
        {
            ("What is Aspose.Words?", "Aspose.Words is a .NET library for creating, editing, converting, and rendering Word documents without Microsoft Word."),
            ("Which file formats are supported?", "DOC, DOCX, ODT, RTF, HTML, PDF, XPS, EPUB and many more."),
            ("Can I render documents to PDF?", "Yes, using the PdfSaveOptions class you can save a document as PDF."),
            ("How do I update fields programmatically?", "Call Document.UpdateFields() before saving the document.")
        };

        // Insert each FAQ entry.
        foreach (var faq in faqs)
        {
            // Question as a heading.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln(faq.Question);

            // Answer as normal paragraph.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln(faq.Answer);
            builder.Writeln(); // Add spacing between entries.
        }

        // Ensure any fields are up‑to‑date.
        doc.UpdateFields();

        // Save the document.
        doc.Save("FAQ.docx");
    }
}
