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

        // ---------- Outer Table ----------
        // Start the outer table.
        Table outerTable = builder.StartTable();

        // First row of the outer table.
        builder.InsertCell();
        builder.Write("Outer Cell 1");
        builder.InsertCell();
        builder.Write("Outer Cell 2");
        builder.EndRow();

        // Second row of the outer table.
        builder.InsertCell();

        // ---------- Nested (Inner) Table ----------
        // Start a nested table inside the first cell of the second row.
        Table innerTable = builder.StartTable();

        // First row of the inner table.
        builder.InsertCell();
        builder.Write("Inner 1");
        builder.InsertCell();
        builder.Write("Inner 2");
        builder.EndRow();

        // Second row of the inner table.
        builder.InsertCell();
        builder.Write("Inner 3");
        builder.InsertCell();
        builder.Write("Inner 4");
        builder.EndRow();

        // End the inner table.
        builder.EndTable();

        // Continue with the second cell of the outer table's second row.
        builder.InsertCell();
        builder.Write("Outer Cell 4");
        builder.EndRow();

        // End the outer table.
        builder.EndTable();

        // Save the document to disk.
        doc.Save("NestedTable.docx");
    }
}
