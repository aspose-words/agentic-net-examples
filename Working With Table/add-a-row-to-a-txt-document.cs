using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to construct a table with a single row.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin the table.
        builder.StartTable();

        // First cell of the row.
        builder.InsertCell();
        builder.Write("Cell 1");

        // Second cell of the same row.
        builder.InsertCell();
        builder.Write("Cell 2");

        // End the current row (adds the Row to the table).
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document as a plain‑text file.
        TxtSaveOptions txtOptions = new TxtSaveOptions(); // default options
        doc.Save("Output.txt", txtOptions);
    }
}
