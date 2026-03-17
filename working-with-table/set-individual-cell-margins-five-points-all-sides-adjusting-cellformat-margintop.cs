using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // First cell – set individual paddings to 5 points.
        builder.InsertCell();
        builder.CellFormat.TopPadding = 5;
        builder.CellFormat.BottomPadding = 5;
        builder.CellFormat.LeftPadding = 5;
        builder.CellFormat.RightPadding = 5;
        builder.Write("Cell with 5pt padding");

        // Second cell – use SetPaddings method for the same effect.
        builder.InsertCell();
        builder.CellFormat.SetPaddings(5, 5, 5, 5);
        builder.Write("Another cell");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        doc.Save("CellMargins.docx");
    }
}
