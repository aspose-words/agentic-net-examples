using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with sample citation text.
        builder.Writeln("Doe, J. (2023). Example citation for hanging indent.");

        // Retrieve the last paragraph (the one we just added).
        Paragraph paragraph = doc.LastSection.Body.Paragraphs[doc.LastSection.Body.Paragraphs.Count - 1];

        // Set hanging indent to 0.25 inches (0.25 * 72 points = 18 points).
        // In Aspose.Words, a hanging indent is achieved by a negative first‑line indent
        // and a positive left indent of the same magnitude.
        paragraph.ParagraphFormat.FirstLineIndent = -18.0; // outdent first line
        paragraph.ParagraphFormat.LeftIndent = 18.0;      // indent subsequent lines

        // Save the document.
        string outputPath = "CitationWithHangingIndent.docx";
        doc.Save(outputPath);
    }
}
