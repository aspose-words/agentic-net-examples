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

        // ---------- First table ----------
        Table firstTable = builder.StartTable();
        builder.InsertCell();
        builder.Write("First Table - Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("First Table - Row 1, Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("First Table - Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("First Table - Row 2, Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Insert a paragraph between the tables (required for Join to work).
        builder.Writeln();

        // ---------- Second table ----------
        Table secondTable = builder.StartTable();
        builder.InsertCell();
        builder.Write("Second Table - Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Second Table - Row 1, Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Second Table - Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Second Table - Row 2, Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the two tables from the document body.
        Table table1 = doc.FirstSection.Body.Tables[0];
        Table table2 = doc.FirstSection.Body.Tables[1];

        // Append all rows from the second table to the first table.
        while (table2.HasChildNodes)
        {
            // Move the first row of the second table to the end of the first table.
            table1.Rows.Add(table2.FirstRow);
        }

        // Remove the now‑empty second table container.
        table2.Remove();

        // Save the resulting document.
        doc.Save("JoinedTables.docx");
    }
}
