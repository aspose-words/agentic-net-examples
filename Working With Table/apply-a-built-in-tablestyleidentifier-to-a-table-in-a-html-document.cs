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
            // Apply a built‑in table style using its identifier.
            table.StyleIdentifier = StyleIdentifier.TableGrid;

            // Optionally specify which parts of the style are applied.
            table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;
        }

        // Save the document with the applied style.
        doc.Save("output.docx");
    }
}
