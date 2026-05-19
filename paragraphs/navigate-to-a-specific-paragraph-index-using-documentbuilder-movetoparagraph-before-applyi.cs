using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several paragraphs into the document.
        builder.Writeln("Paragraph 1: left aligned.");
        builder.Writeln("Paragraph 2: left aligned.");
        builder.Writeln("Paragraph 3: will be centered.");
        builder.Writeln("Paragraph 4: left aligned.");

        // Move the builder's cursor to the third paragraph (zero‑based index 2).
        // The character index 0 places the cursor at the start of that paragraph.
        builder.MoveToParagraph(2, 0);

        // Apply formatting to the paragraph that the builder is now positioned on.
        builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Add additional text to the same paragraph after applying the formatting.
        builder.Writeln("Additional centered text.");

        // Save the modified document to a file.
        doc.Save("ParagraphNavigationExample.docx");
    }
}
