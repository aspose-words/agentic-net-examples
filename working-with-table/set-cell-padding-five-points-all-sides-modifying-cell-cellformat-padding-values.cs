using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // Apply 5‑point padding to every cell that will be created from now on.
        builder.CellFormat.SetPaddings(5, 5, 5, 5);

        // First cell.
        builder.InsertCell();
        builder.Write("First cell with 5 pt padding.");

        // Second cell.
        builder.InsertCell();
        builder.Write("Second cell with 5 pt padding.");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Directly modify an existing cell's padding (optional, demonstrates the same effect).
        Cell firstCell = table.FirstRow.FirstCell;
        firstCell.CellFormat.LeftPadding = 5;
        firstCell.CellFormat.RightPadding = 5;
        firstCell.CellFormat.TopPadding = 5;
        firstCell.CellFormat.BottomPadding = 5;

        // Save the document.
        doc.Save("CellPaddingExample.docx");
    }
}
