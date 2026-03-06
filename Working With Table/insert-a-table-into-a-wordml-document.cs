using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeTableInsertExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new empty document.
            Document doc = new Document();

            // Initialize a DocumentBuilder which simplifies node insertion.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table. The method returns the created Table node.
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

            // Optionally, set a title and description for the table.
            table.Title = "Sample Table Title";
            table.Description = "Sample Table Description";

            // Save the document to a file.
            doc.Save("TableInsert.docx");
        }
    }
}
