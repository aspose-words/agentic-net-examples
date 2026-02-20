using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a single cell.
        Table table = builder.StartTable();
        builder.InsertCell();

        // Insert the desired text into the current cell.
        builder.Write("Hello from cell");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Set save options to preserve the table layout when exporting to plain text.
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        saveOptions.PreserveTableLayout = true;

        // Save the document as a TXT file.
        doc.Save("CellInTxt.txt", saveOptions);
    }
}
