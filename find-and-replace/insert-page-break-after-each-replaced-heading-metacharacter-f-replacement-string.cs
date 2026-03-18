using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class InsertPageBreakAfterHeading
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample headings that we will later replace.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter One");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.2");

        // Replace each heading text with the same text followed by the page‑break meta‑character "\f".
        // In a replacement string the backslash must be escaped, therefore we use "\\f".
        // The Range.Replace method processes meta‑characters, so "\f" will become a page break.
        doc.Range.Replace("Chapter One", "Chapter One\\f");
        doc.Range.Replace("Section 1.1", "Section 1.1\\f");
        doc.Range.Replace("Section 1.2", "Section 1.2\\f");

        // Save the document to disk.
        doc.Save("OutputWithPageBreaks.docx");
    }
}
