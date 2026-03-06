using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the HTML document.
        Document doc = new Document("input.html");

        // Locate the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table != null)
        {
            // Apply a built‑in table style identifier (e.g., TableGrid).
            table.StyleIdentifier = StyleIdentifier.TableGrid;

            // Optionally specify which parts of the style to apply.
            table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;
        }

        // Save the modified document.
        doc.Save("output.docx");
    }
}
