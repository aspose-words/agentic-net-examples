using System;
using Aspose.Words;
using Aspose.Words.Tables;

class DeleteThirdColumn
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with at least three columns.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell(); builder.Writeln("Row 1, Cell 1");
        builder.InsertCell(); builder.Writeln("Row 1, Cell 2");
        builder.InsertCell(); builder.Writeln("Row 1, Cell 3");
        builder.EndRow();

        // Second row.
        builder.InsertCell(); builder.Writeln("Row 2, Cell 1");
        builder.InsertCell(); builder.Writeln("Row 2, Cell 2");
        builder.InsertCell(); builder.Writeln("Row 2, Cell 3");
        builder.EndRow();

        builder.EndTable();

        // Remove the third column (index 2) from each row, if it exists.
        foreach (Row row in table.Rows)
        {
            if (row.Cells.Count > 2)
                row.Cells.RemoveAt(2);
        }

        // Save the modified document.
        doc.Save("output.docx");
    }
}
