using System;
using Aspose.Words;
using Aspose.Words.Tables;

class JoinTablesExample
{
    static void Main()
    {
        // Load the document that already contains two tables.
        string inputPath = @"C:\Docs\Input.docx";
        Document doc = new Document(inputPath);

        // Access the collection of tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Ensure there are at least two tables to join.
        if (tables.Count < 2)
        {
            Console.WriteLine("The document must contain at least two tables.");
            return;
        }

        // Get references to the two tables we want to join.
        Table firstTable = tables[0];
        Table secondTable = tables[1];

        // Append a copy of each row from the second table to the first table.
        // Clone each row (deep clone) so that the original second table can be removed safely.
        foreach (Row row in secondTable.Rows)
        {
            Row clonedRow = (Row)row.Clone(true);
            firstTable.Rows.Add(clonedRow);
        }

        // Remove the now‑empty second table from the document.
        secondTable.Remove();

        // Save the modified document.
        string outputPath = @"C:\Docs\Output.docx";
        doc.Save(outputPath);
    }
}
