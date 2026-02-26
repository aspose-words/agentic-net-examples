using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the source PDF document that already contains at least two tables.
        Document doc = new Document("input.pdf");

        // Retrieve the first table from the document's body.
        Table firstTable = doc.FirstSection.Body.Tables[0];

        // Retrieve the second table using the generic GetChild method.
        Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

        // Transfer all rows from the second table to the first table.
        while (secondTable.HasChildNodes)
        {
            // Add the first row of the second table to the first table.
            // This also removes the row from the second table.
            firstTable.Rows.Add(secondTable.FirstRow);
        }

        // Remove the now‑empty second table container from the document.
        secondTable.Remove();

        // Save the modified document back to PDF format.
        doc.Save("output.pdf", SaveFormat.Pdf);
    }
}
