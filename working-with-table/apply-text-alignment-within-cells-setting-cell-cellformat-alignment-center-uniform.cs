using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin a table.
        builder.StartTable();

        // ---- First row ----
        // Cell 1
        builder.InsertCell();
        // Center-align the paragraph text inside this cell.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Cell 1");

        // Cell 2
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Cell 2");

        // Finish the first row.
        builder.EndRow();

        // ---- Second row ----
        // Cell 3
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Cell 3");

        // Cell 4
        builder.InsertCell();
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Write("Cell 4");

        // Finish the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to disk.
        doc.Save("AlignedTable.docx");
    }
}
