using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing RTF document.
        Document doc = new Document("Input.rtf");

        // Retrieve the first two tables in the document.
        Table table1 = doc.FirstSection.Body.Tables[0];
        Table table2 = doc.FirstSection.Body.Tables[1];

        // Append all rows from the second table to the first table.
        // Clone each row to keep the original table intact until removal.
        foreach (Row row in table2.Rows)
        {
            Row clonedRow = (Row)row.Clone(true);
            table1.Rows.Add(clonedRow);
        }

        // Remove the now redundant second table from the document.
        table2.Remove();

        // Save the modified document back to RTF format.
        doc.Save("Output.rtf");
    }
}
