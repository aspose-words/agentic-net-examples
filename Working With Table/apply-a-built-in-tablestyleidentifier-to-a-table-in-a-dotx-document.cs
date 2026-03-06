using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load an existing DOTX template (or create a new document if you prefer).
        Document doc = new Document("Template.dotx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin building a table.
        Table table = builder.StartTable();

        // Insert the first row (required before setting any table formatting).
        builder.InsertCell();
        builder.Writeln("Header");
        builder.EndRow();

        // Insert a second row with sample data.
        builder.InsertCell();
        builder.Writeln("Data");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply a built‑in table style using its identifier.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Optionally specify which parts of the style to apply.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Adjust column widths to fit the content.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document back to DOTX (or any other supported format).
        doc.Save("StyledTable.dotx");
    }
}
