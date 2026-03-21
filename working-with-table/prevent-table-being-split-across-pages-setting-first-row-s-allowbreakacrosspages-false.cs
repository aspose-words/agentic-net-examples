using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableRowBreakControl
{
    class Program
    {
        static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a simple table with two rows.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("First row, first cell");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln("Second row, first cell");
            builder.EndRow();
            builder.EndTable();

            // Access the first row of the table.
            Row firstRow = table.FirstRow;

            // Prevent this row from being split across pages.
            firstRow.RowFormat.AllowBreakAcrossPages = false;

            // Save the modified document.
            doc.Save("Output.docx");
        }
    }
}
