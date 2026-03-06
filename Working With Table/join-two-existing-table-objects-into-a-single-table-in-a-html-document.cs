using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the HTML document that contains the tables.
        Document doc = new Document("input.html");

        // Retrieve the first two tables from the document body.
        Table firstTable = doc.FirstSection.Body.Tables[0];
        Table secondTable = doc.FirstSection.Body.Tables[1];

        // Transfer all rows from the second table to the first table.
        while (secondTable.HasChildNodes)
        {
            // Add the first row of the second table to the first table.
            // This also removes the row from the second table.
            firstTable.Rows.Add(secondTable.FirstRow);
        }

        // Remove the now‑empty second table container.
        secondTable.Remove();

        // Save the modified document back to HTML.
        doc.Save("output.html");
    }
}
