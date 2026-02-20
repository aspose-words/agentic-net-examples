using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyle
{
    static void Main()
    {
        // Load the DOTX template.
        Document doc = new Document("Template.dotx");

        // Get the first table in the document (ensure a table exists).
        Table table = doc.FirstSection.Body.Tables[0];

        // Apply a built‑in table style using its identifier.
        table.StyleIdentifier = StyleIdentifier.LightGrid;

        // Optionally specify which parts of the style are applied.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Save the document with the applied style.
        doc.Save("StyledTable.docx");
    }
}
