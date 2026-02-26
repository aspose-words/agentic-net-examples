using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table. The first call to InsertCell will also start the first row.
        builder.StartTable();

        // Insert a new cell into the current row and obtain the Cell object.
        Cell cell = builder.InsertCell();

        // Write the desired text into the current cell.
        builder.Write("This is the text inside the cell.");

        // End the current row (optional here since we have only one row).
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to a file.
        doc.Save("CellText.docx");
    }
}
