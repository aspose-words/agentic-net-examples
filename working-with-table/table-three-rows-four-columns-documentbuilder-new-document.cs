using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table. The method returns the Table node.
        Table table = builder.StartTable();

        // Build three rows.
        for (int row = 0; row < 3; row++)
        {
            // Build four columns (cells) in each row.
            for (int col = 0; col < 4; col++)
            {
                // Insert a new cell into the current row.
                builder.InsertCell();

                // Write some text into the cell.
                builder.Write($"Row {row + 1}, Col {col + 1}");
            }

            // End the current row before starting the next one.
            builder.EndRow();
        }

        // End the table construction.
        builder.EndTable();

        // Save the document to a file.
        doc.Save("Table3x4.docx");
    }
}
