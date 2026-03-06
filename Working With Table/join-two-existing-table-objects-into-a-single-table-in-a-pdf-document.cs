using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the PDF document that contains the tables.
        Document doc = new Document("input.pdf");

        // Retrieve the first two tables in the document.
        Table firstTable = doc.FirstSection.Body.Tables[0];
        Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

        // Transfer all rows from the second table to the first table.
        while (secondTable.HasChildNodes)
        {
            // Add the current first row of the second table to the first table.
            firstTable.Rows.Add(secondTable.FirstRow);
        }

        // Remove the now empty second table container.
        secondTable.Remove();

        // Save the modified document as a PDF.
        doc.Save("output.pdf");
    }
}
