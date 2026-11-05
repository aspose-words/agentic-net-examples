using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the outer table (2 rows x 2 columns).
        Table outerTable = builder.StartTable();

        // ----- First row of outer table -----
        // Cell (1,1)
        builder.InsertCell();
        builder.Writeln("Outer Cell 1,1");

        // Cell (1,2) - this cell will contain a nested table.
        builder.InsertCell();

        // Start the nested table inside the current outer cell.
        Table nestedTable = builder.StartTable();

        // First row of nested table.
        builder.InsertCell();
        builder.Writeln("Nested Cell 1,1");
        builder.InsertCell();
        builder.Writeln("Nested Cell 1,2");
        builder.EndRow();

        // Second row of nested table.
        builder.InsertCell();
        builder.Writeln("Nested Cell 2,1");
        builder.InsertCell();
        builder.Writeln("Nested Cell 2,2");
        builder.EndRow();

        // End the nested table.
        builder.EndTable();

        // End the first row of the outer table.
        builder.EndRow();

        // ----- Second row of outer table -----
        // Cell (2,1)
        builder.InsertCell();
        builder.Writeln("Outer Cell 2,1");

        // Cell (2,2)
        builder.InsertCell();
        builder.Writeln("Outer Cell 2,2");
        builder.EndRow();

        // End the outer table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "NestedTable.docx";
        doc.Save(outputPath);
    }
}
