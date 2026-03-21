using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document and add a simple table with one row.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.StartTable();
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Cell {i + 1}");
        }
        builder.EndRow();
        builder.EndTable();

        // Retrieve the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Determine how many columns the table has (assumes a uniform layout).
        int columnCount = table.FirstRow.Cells.Count;

        // Create a new row that belongs to the same document.
        Row newRow = new Row(doc);

        // Populate the new row with cells containing placeholder text.
        for (int i = 0; i < columnCount; i++)
        {
            // Create a new cell.
            Cell cell = new Cell(doc);

            // Ensure the cell has at least one paragraph.
            cell.EnsureMinimum();

            // Insert placeholder text into the cell's first paragraph.
            cell.FirstParagraph.AppendChild(new Run(doc, $"Placeholder {i + 1}"));

            // Add the cell to the new row.
            newRow.AppendChild(cell);
        }

        // Append the completed row to the end of the table.
        table.AppendChild(newRow);

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
