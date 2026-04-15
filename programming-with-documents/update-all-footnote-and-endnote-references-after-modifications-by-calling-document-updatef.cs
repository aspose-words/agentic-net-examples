using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text and insert a footnote.
        builder.Write("This is a paragraph with a footnote");
        Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "Original footnote text.");

        // Insert an endnote after the footnote.
        builder.Write(" and an endnote");
        Footnote endnote = builder.InsertFootnote(FootnoteType.Endnote, "Original endnote text.");

        // Modify the footnote text to simulate a document change.
        builder.MoveTo(footnote.FirstParagraph);
        builder.Write(" – updated footnote content.");

        // Update all fields in the document (required for correct reference marks).
        doc.UpdateFields();

        // Update the actual reference marks of footnotes and endnotes.
        doc.UpdateActualReferenceMarks();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UpdatedFootnotes.docx");
        doc.Save(outputPath);
    }
}
