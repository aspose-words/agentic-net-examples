using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Define the folder where the output document will be saved.
        string artifactsDir = @"C:\Output\"; // Adjust the path as needed.

        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder which provides a convenient API for building the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table. The method returns the Table node that has been created.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to the specified location.
        doc.Save(artifactsDir + "InsertedTable.docx");
    }
}
