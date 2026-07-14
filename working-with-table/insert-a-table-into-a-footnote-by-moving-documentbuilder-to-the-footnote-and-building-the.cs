using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Notes;   // Required for Footnote and FootnoteType

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text to the main body.
        builder.Write("This is a paragraph with a footnote reference.");

        // Insert a footnote and obtain the Footnote node.
        Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.");

        // Move the builder cursor into the footnote's first paragraph.
        builder.MoveTo(footnote.FirstParagraph);

        // Build a table inside the footnote.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Optionally move the cursor back to the end of the document.
        builder.MoveToDocumentEnd();

        // Save the resulting document.
        doc.Save("FootnoteTable.docx");
    }
}
