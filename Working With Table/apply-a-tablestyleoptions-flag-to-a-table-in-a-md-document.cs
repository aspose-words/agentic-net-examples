using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the Markdown document that contains a table.
        Document doc = new Document("input.md");

        // Retrieve the first table in the document (if any).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table != null)
        {
            // Apply the desired TableStyleOptions flags.
            // Example: apply formatting to the first row and enable row banding.
            table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;
        }

        // Save the modified document. The output format can be any supported type (e.g., DOCX).
        doc.Save("output.docx");
    }
}
