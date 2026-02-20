using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing DOTM document.
        Document doc = new Document("Input.dotm");

        // Check that the document contains at least one table.
        if (doc.FirstSection.Body.Tables.Count > 0)
        {
            // Get the first table in the document.
            Table table = doc.FirstSection.Body.Tables[0];

            // (Optional) Assign a built‑in style to the table.
            table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

            // Apply the desired TableStyleOptions flags.
            table.StyleOptions = TableStyleOptions.FirstRow |
                                 TableStyleOptions.RowBands |
                                 TableStyleOptions.FirstColumn;
        }

        // Save the modified document back to DOTM format.
        doc.Save("Output.dotm");
    }
}
