using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Instantiate a Table node and attach it to the document body.
        Table table = new Table(doc);
        doc.FirstSection.Body.AppendChild(table);

        // Create a Row node and add it to the table.
        Row row = new Row(doc);
        table.AppendChild(row);

        // Create a Cell node and add it to the row.
        Cell cell = new Cell(doc);
        row.AppendChild(cell);

        // Every cell must contain at least one Paragraph.
        Paragraph paragraph = new Paragraph(doc);
        cell.AppendChild(paragraph);

        // Add a Run with some text to the paragraph.
        Run run = new Run(doc, "Hello world!");
        paragraph.AppendChild(run);

        // Save the document to disk.
        doc.Save("Table_CreateTable.docx");
    }
}
