using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

namespace AsposeWordsFootnoteDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph with a footnote.
            builder.Writeln("This is a paragraph that contains a footnote.");
            builder.InsertFootnote(FootnoteType.Footnote, "This is the footnote text.");

            // Insert another paragraph with an endnote.
            builder.Writeln("This is a paragraph that contains an endnote.");
            builder.InsertFootnote(FootnoteType.Endnote, "This is the endnote text.");

            // Make an additional modification to demonstrate that references need updating.
            builder.Writeln("Additional content added after footnotes and endnotes.");

            // Update all fields in the document (required for correct footnote/endnote references).
            doc.UpdateFields();

            // Update the actual reference marks of footnotes and endnotes.
            doc.UpdateActualReferenceMarks();

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string outputPath = Path.Combine(outputDir, "FootnoteEndnoteUpdated.docx");
            doc.Save(outputPath);
        }
    }
}
