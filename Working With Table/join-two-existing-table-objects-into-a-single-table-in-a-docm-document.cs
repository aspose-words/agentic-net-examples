using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing DOCM document.
        Document doc = new Document("input.docm");

        // Retrieve the first two tables in the document.
        Table firstTable = doc.FirstSection.Body.Tables[0];
        Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

        // Transfer all rows from the second table to the first table.
        while (secondTable.HasChildNodes)
        {
            firstTable.Rows.Add(secondTable.FirstRow);
        }

        // Remove the now‑empty second table container.
        secondTable.Remove();

        // Save the document with the combined table.
        doc.Save("output.docm");
    }
}
