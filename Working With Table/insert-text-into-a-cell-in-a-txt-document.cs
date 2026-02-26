using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 1x1 table.
        builder.StartTable();          // Begin the table.
        builder.InsertCell();          // Insert the first (and only) cell.
        builder.Write("Text inside the cell."); // Insert the desired text into the cell.
        builder.EndRow();              // End the row.
        builder.EndTable();            // End the table.

        // Save the document as plain‑text, preserving the table layout.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            PreserveTableLayout = true
        };
        doc.Save("CellInTxt.txt", txtOptions);
    }
}
