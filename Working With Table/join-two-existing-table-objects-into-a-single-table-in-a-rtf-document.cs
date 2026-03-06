using Aspose.Words;
using Aspose.Words.Tables;

class TableJoinExample
{
    static void Main()
    {
        // Load the RTF document that contains the two tables to be combined.
        Document doc = new Document("Input.rtf");

        // Retrieve the first table from the document's body collection.
        Table firstTable = doc.FirstSection.Body.Tables[0];

        // Retrieve the second table using the generic GetChild method.
        Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

        // Transfer all rows from the second table to the first table.
        while (secondTable.HasChildNodes)
        {
            // Append the current first row of the second table to the first table.
            firstTable.Rows.Add(secondTable.FirstRow);
        }

        // Remove the now empty second table from the document.
        secondTable.Remove();

        // Save the updated document back to RTF format.
        doc.Save("Output.rtf", SaveFormat.Rtf);
    }
}
