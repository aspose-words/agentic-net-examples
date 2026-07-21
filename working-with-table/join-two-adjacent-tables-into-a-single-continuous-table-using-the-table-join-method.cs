using System;
using System.IO;
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
        // Build a 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Table1 Row1 Cell1");
        builder.InsertCell();
        builder.Write("Table1 Row1 Cell2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Table1 Row2 Cell1");
        builder.InsertCell();
        builder.Write("Table1 Row2 Cell2");
        builder.EndTable();

        // ---------- Second table ----------
        // Build a 1x2 table directly after the first one (no paragraph between them).
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Table2 Row1 Cell1");
        builder.InsertCell();
        builder.Write("Table2 Row1 Cell2");
        builder.EndTable();

        // Retrieve the two tables from the document body.
        Table firstTable = doc.FirstSection.Body.Tables[0];
        Table secondTable = doc.FirstSection.Body.Tables[1];

        // Join the second table into the first one by moving all rows.
        while (secondTable.HasChildNodes)
        {
            // Move the first row of the second table to the end of the first table.
            firstTable.Rows.Add(secondTable.FirstRow);
        }

        // Remove the now‑empty second table container.
        secondTable.Remove();

        // Validate that the tables were joined correctly.
        // The first table should now contain 3 rows (2 + 1).
        int expectedRows = 3;
        if (firstTable.Rows.Count != expectedRows)
            throw new InvalidOperationException(
                $"Table join failed. Expected {expectedRows} rows, but found {firstTable.Rows.Count}.");

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "JoinedTables.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
