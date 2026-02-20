using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing DOTM template.
        Document doc = new Document("Template.dotm");

        // Get the first table in the document (adjust index if needed).
        Table table = doc.FirstSection.Body.Tables[0];

        // Create a new row that belongs to the same document.
        Row newRow = new Row(doc);

        // Create the first cell and add some text.
        Cell cell1 = new Cell(doc);
        cell1.FirstParagraph.AppendChild(new Run(doc, "New Cell 1"));
        newRow.Cells.Add(cell1);

        // Create the second cell and add some text.
        Cell cell2 = new Cell(doc);
        cell2.FirstParagraph.AppendChild(new Run(doc, "New Cell 2"));
        newRow.Cells.Add(cell2);

        // Append the new row to the end of the table.
        table.Rows.Add(newRow);

        // Save the modified document back to DOTM format.
        doc.Save("Result.dotm");
    }
}
