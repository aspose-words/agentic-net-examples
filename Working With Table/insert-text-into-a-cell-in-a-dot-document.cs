using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and insert the first cell.
        Table table = builder.StartTable();
        builder.InsertCell();

        // Insert the desired text into the current cell.
        builder.Write("Hello from the cell!");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document as a DOT template.
        doc.Save("Output.dot", SaveFormat.Dot);
    }
}
