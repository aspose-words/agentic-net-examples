using System;
using Aspose.Words;
using Aspose.Words.Tables;

class JoinTablesExample
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Access the first two tables in the document.
        // Adjust the indices if the tables are located elsewhere.
        Table firstTable = doc.FirstSection.Body.Tables[0];
        Table secondTable = doc.FirstSection.Body.Tables[1];

        // Append all rows from the second table to the first table.
        // Clone each row to keep the original table intact until removal.
        foreach (Row row in secondTable.Rows)
        {
            Row clonedRow = (Row)row.Clone(true);
            firstTable.Rows.Add(clonedRow);
        }

        // Remove the now redundant second table from the document.
        secondTable.Remove();

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
