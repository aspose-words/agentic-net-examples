using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the document that contains the two tables to be merged.
        Document doc = new Document("Input.docx");

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

        // Remove the now empty second table container from the document.
        secondTable.Remove();

        // Save the resulting document with the combined table.
        doc.Save("Output.docx");
    }
}
