using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a hanging indent of 0.25 inches (0.25 * 72 = 18 points) by using a negative FirstLineIndent.
        builder.ParagraphFormat.FirstLineIndent = -18; // points

        // Optional: set alignment or other formatting as needed for citations.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

        // Write the citation text.
        builder.Writeln("Doe, J. (2023). Example citation for a reference list.");

        // Save the document.
        doc.Save("CitationWithHangingIndent.docx");
    }
}
