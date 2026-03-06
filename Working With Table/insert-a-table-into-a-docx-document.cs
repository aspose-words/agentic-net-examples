using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Associate a DocumentBuilder with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table; the method returns the created Table node.
        Table table = builder.StartTable();

        // ---- First row ----
        builder.InsertCell();               // First cell of the first row.
        builder.Write("Row 1, Cell 1.");    // Add text to the cell.

        builder.InsertCell();               // Second cell of the first row.
        builder.Write("Row 1, Cell 2.");    // Add text to the cell.

        builder.EndRow();                   // Finish the first row.

        // ---- Second row ----
        builder.InsertCell();               // First cell of the second row.
        builder.Write("Row 2, Cell 1.");    // Add text to the cell.

        builder.InsertCell();               // Second cell of the second row.
        builder.Write("Row 2, Cell 2.");    // Add text to the cell.

        builder.EndRow();                   // Finish the second row.

        // End the table construction.
        builder.EndTable();

        // Save the document as a DOCX file.
        doc.Save("TableExample.docx");
    }
}
