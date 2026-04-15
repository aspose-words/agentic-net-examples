using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several paragraphs to have content to navigate.
        builder.Writeln("Paragraph 0: First paragraph.");
        builder.Writeln("Paragraph 1: Second paragraph.");
        builder.Writeln("Paragraph 2: Third paragraph.");
        builder.Writeln("Paragraph 3: Fourth paragraph.");

        // Move the builder's cursor to the third paragraph (zero‑based index 2) at the start of the paragraph.
        builder.MoveToParagraph(2, 0);

        // Apply formatting changes to the paragraph at the current cursor position.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Font.Bold = true;
        builder.Font.Size = 14;
        builder.Writeln("This paragraph has been formatted.");

        // Save the modified document to disk.
        doc.Save("FormattedParagraph.docx");
    }
}
