using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some main text.
        builder.Write("This is a paragraph with a footnote reference.");

        // Insert a footnote.
        Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.");

        // Move the builder into the footnote's paragraph.
        builder.MoveTo(footnote.FirstParagraph);

        // Start a table inside the footnote.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1, Row 1");
        builder.InsertCell();
        builder.Write("Cell 2, Row 1");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 1, Row 2");
        builder.InsertCell();
        builder.Write("Cell 2, Row 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Add additional text after the table within the footnote.
        builder.Writeln();
        builder.Write("Additional footnote text after the table.");

        // Save the document.
        doc.Save("FootnoteTable.docx");
    }
}
