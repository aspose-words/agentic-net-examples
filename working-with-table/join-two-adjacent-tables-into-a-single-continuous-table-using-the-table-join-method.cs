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

        // Build the first table (2 rows, 2 columns).
        Table firstTable = builder.StartTable();
        builder.InsertCell();
        builder.Write("A1");
        builder.InsertCell();
        builder.Write("B1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("A2");
        builder.InsertCell();
        builder.Write("B2");
        builder.EndTable();

        // Build the second table (2 rows, 2 columns) directly after the first one.
        Table secondTable = builder.StartTable();
        builder.InsertCell();
        builder.Write("C1");
        builder.InsertCell();
        builder.Write("D1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("C2");
        builder.InsertCell();
        builder.Write("D2");
        builder.EndTable();

        // Combine the second table into the first table by moving all rows.
        while (secondTable.HasChildNodes)
        {
            // Move the first row of the second table to the end of the first table.
            firstTable.Rows.Add(secondTable.FirstRow);
        }

        // Remove the now‑empty second table container.
        secondTable.Remove();

        // Verify that only one table remains.
        if (doc.FirstSection.Body.Tables.Count != 1)
            throw new InvalidOperationException("Table join failed: more than one table present.");

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "JoinedTables.docx");
        doc.Save(outputPath);

        // Simple existence check (throws if the file was not created).
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output file was not created.", outputPath);
    }
}
