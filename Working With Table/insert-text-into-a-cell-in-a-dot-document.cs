using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTextIntoCell
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table. The first call to InsertCell also creates the first row.
        builder.StartTable();

        // Insert the first cell of the first row.
        Cell cell = builder.InsertCell();

        // Write the desired text into the current cell.
        builder.Write("Hello, this is text inside the cell.");

        // End the row and the table (optional if only one cell is needed).
        builder.EndRow();
        builder.EndTable();

        // Save the document as a DOT template (Word 97-2003 template).
        doc.Save("OutputTemplate.dot");
    }
}
