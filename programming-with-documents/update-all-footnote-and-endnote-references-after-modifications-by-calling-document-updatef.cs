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
        builder.Writeln("This paragraph contains a footnote.");
        Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "Initial footnote text.");

        // Insert a paragraph with an endnote.
        builder.Writeln("This paragraph contains an endnote.");
        Footnote endnote = builder.InsertFootnote(FootnoteType.Endnote, "Initial endnote text.");

        // Modify the footnote text.
        builder.MoveTo(footnote.FirstParagraph);
        builder.Write(" Updated footnote content.");

        // Modify the endnote text.
        builder.MoveTo(endnote.FirstParagraph);
        builder.Write(" Updated endnote content.");

        // Update all fields in the document (if any) and then refresh footnote/endnote reference marks.
        doc.UpdateFields();
        doc.UpdateActualReferenceMarks();

        // Save the resulting document.
        doc.Save("FootnoteEndnoteUpdated.docx");
    }
}
