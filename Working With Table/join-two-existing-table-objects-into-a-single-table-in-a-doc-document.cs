using System;
using Aspose.Words;
using Aspose.Words.Tables;

class JoinTablesExample
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("Input.docx");

        // Assume the document contains at least two tables.
        Table firstTable = doc.FirstSection.Body.Tables[0];
        Table secondTable = doc.FirstSection.Body.Tables[1];

        // Append all rows from the second table to the first table.
        // Clone each row to keep the original table intact while moving the content.
        foreach (Row row in secondTable.Rows)
        {
            // Deep clone the row (including its cells and contents).
            Row clonedRow = (Row)row.Clone(true);
            firstTable.Rows.Add(clonedRow);
        }

        // Remove the now redundant second table from the document.
        secondTable.Remove();

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
