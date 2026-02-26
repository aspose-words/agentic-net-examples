using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();

        // Ensure the document has a body to work with.
        doc.FirstSection.Body.EnsureMinimum();

        // Create a table and add it to the document.
        Table table = new Table(doc);
        doc.FirstSection.Body.AppendChild(table);

        // Create a new row and append it to the table.
        Row newRow = new Row(doc);
        table.AppendChild(newRow);

        // Ensure the row contains at least one cell.
        newRow.EnsureMinimum();

        // Add some text to the first cell of the new row.
        newRow.FirstCell.FirstParagraph.AppendChild(new Run(doc, "New row added"));

        // Save the document.
        doc.Save("AddedRow.docx");
    }
}
