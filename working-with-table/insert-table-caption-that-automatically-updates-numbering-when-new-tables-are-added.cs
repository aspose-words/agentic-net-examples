using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        builder.EndTable();

        // Insert a caption paragraph directly after the table.
        // The caption consists of the label "Table", a SEQ field that provides automatic numbering,
        // and the caption text. The SEQ field will be updated whenever fields are updated.
        builder.Writeln(); // Move to a new paragraph.
        builder.Write("Table ");
        // SEQ field for table numbering; the result placeholder ("1") will be replaced after UpdateFields().
        builder.InsertField("SEQ Table \\* ARABIC", "1");
        builder.Write(" – Sample table caption.");

        // Update all fields so the caption shows the correct number.
        doc.UpdateFields();

        // Save the document.
        doc.Save("TableWithCaption.docx");
    }
}
