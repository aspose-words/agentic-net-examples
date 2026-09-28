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

        // Start the outer table.
        Table outerTable = builder.StartTable();

        // First cell of the outer table.
        builder.InsertCell();
        builder.Writeln("Outer Cell 1");

        // Second cell of the outer table – this will contain the nested table.
        builder.InsertCell();

        // Start the nested table inside the current cell.
        Table nestedTable = builder.StartTable();

        // First row of the nested table.
        builder.InsertCell();
        builder.Writeln("Nested Cell 1");
        builder.InsertCell();
        builder.Writeln("Nested Cell 2");
        builder.EndRow();

        // Second row of the nested table.
        builder.InsertCell();
        builder.Writeln("Nested Cell 3");
        builder.InsertCell();
        builder.Writeln("Nested Cell 4");
        builder.EndRow();

        // End the nested table.
        builder.EndTable();

        // End the row of the outer table.
        builder.EndRow();

        // End the outer table.
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "NestedTable.docx";
        doc.Save(outputPath);
    }
}
