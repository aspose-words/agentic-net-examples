using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Notes;

class Program
{
    static void Main()
    {
        // Create a new empty document and associate a DocumentBuilder with it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a footnote. The footnote text is left empty because we will fill it with a table.
        Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, string.Empty);

        // Move the builder's cursor to the first paragraph of the footnote so that subsequent inserts go inside the footnote.
        builder.MoveTo(footnote.FirstParagraph);

        // Start building a table inside the footnote.
        Table table = builder.StartTable();

        // ---- First row ----
        builder.InsertCell();               // First cell of the row.
        builder.Write("Cell 1,1");          // Write text into the first cell.
        builder.InsertCell();               // Second cell of the row.
        builder.Write("Cell 1,2");          // Write text into the second cell.
        builder.EndRow();                   // End the first row.

        // ---- Second row ----
        builder.InsertCell();               // First cell of the second row.
        builder.Write("Cell 2,1");
        builder.InsertCell();               // Second cell of the second row.
        builder.Write("Cell 2,2");
        builder.EndRow();                   // End the second row.

        // Finish the table. The cursor now points just after the table inside the footnote.
        builder.EndTable();

        // Optionally add more text after the table within the same footnote.
        builder.Writeln();                  // Insert a paragraph break.
        builder.Write("Additional footnote text.");

        // Save the resulting document.
        doc.Save("FootnoteWithTable.docx");
    }
}
