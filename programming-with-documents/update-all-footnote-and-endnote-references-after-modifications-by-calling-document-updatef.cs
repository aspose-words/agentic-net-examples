using System;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with a footnote.
        builder.Writeln("This is some text with a footnote.");
        Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "Original footnote text.");

        // Insert a paragraph with an endnote.
        builder.Writeln("This is some text with an endnote.");
        Footnote endnote = builder.InsertFootnote(FootnoteType.Endnote, "Original endnote text.");

        // Modify the footnote text.
        builder.MoveTo(footnote.FirstParagraph);
        builder.Write(" Updated footnote text.");

        // Modify the endnote text.
        builder.MoveTo(endnote.FirstParagraph);
        builder.Write(" Updated endnote text.");

        // Update fields in the document.
        doc.UpdateFields();

        // Update the actual reference marks for footnotes and endnotes.
        doc.UpdateActualReferenceMarks();

        // Save the document.
        doc.Save("UpdatedFootnotes.docx");
    }
}
