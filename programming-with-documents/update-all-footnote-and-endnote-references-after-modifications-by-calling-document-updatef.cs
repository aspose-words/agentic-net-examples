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
        builder.Writeln("Paragraph with a footnote.");
        builder.InsertFootnote(FootnoteType.Footnote, "Original footnote text.");

        // Insert a paragraph with an endnote.
        builder.Writeln("Paragraph with an endnote.");
        builder.InsertFootnote(FootnoteType.Endnote, "Original endnote text.");

        // Modify the footnote text to simulate a change that requires reference updates.
        Footnote footnote = (Footnote)doc.GetChild(NodeType.Footnote, 0, true);
        footnote.FirstParagraph.AppendChild(new Run(doc, " Additional content."));

        // Update all fields in the document (necessary for many field types).
        doc.UpdateFields();

        // Update the actual reference marks of footnotes and endnotes.
        doc.UpdateActualReferenceMarks();

        // Save the resulting document.
        const string outputFile = "UpdatedFootnotes.docx";
        doc.Save(outputFile);
    }
}
