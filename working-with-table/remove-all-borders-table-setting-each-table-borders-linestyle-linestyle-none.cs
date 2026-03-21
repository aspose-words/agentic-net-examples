using System;
using Aspose.Words;
using Aspose.Words.Tables;

class RemoveTableBorders
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with two cells.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Get the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Remove all borders from the table.
        table.ClearBorders();

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
