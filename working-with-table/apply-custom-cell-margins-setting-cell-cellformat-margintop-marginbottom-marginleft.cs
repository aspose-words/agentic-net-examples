using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        builder.StartTable();

        // Insert the first cell and keep a reference to it.
        Cell firstCell = builder.InsertCell();
        builder.Write("Cell with custom margins.");

        // Apply custom margins (implemented as padding) to the cell.
        firstCell.CellFormat.TopPadding = 20;    // Top margin in points.
        firstCell.CellFormat.BottomPadding = 20; // Bottom margin in points.
        firstCell.CellFormat.LeftPadding = 15;   // Left margin in points.
        firstCell.CellFormat.RightPadding = 15;  // Right margin in points.

        // Insert a second cell without custom margins.
        builder.InsertCell();
        builder.Write("Regular cell.");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to disk.
        doc.Save("CustomCellMargins.docx");
    }
}
