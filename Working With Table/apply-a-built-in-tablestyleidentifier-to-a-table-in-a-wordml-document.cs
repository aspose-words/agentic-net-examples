using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin a new table.
        Table table = builder.StartTable();

        // Insert at least one cell before applying any table formatting.
        builder.InsertCell();
        builder.Writeln("Sample cell");

        // Apply a built‑in table style (e.g., TableGrid).
        table.StyleIdentifier = StyleIdentifier.TableGrid;

        // Optionally specify which parts of the style are applied.
        table.StyleOptions = TableStyleOptions.FirstRow |
                             TableStyleOptions.FirstColumn |
                             TableStyleOptions.RowBands;

        // Resize the table to fit its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Finish the table.
        builder.EndTable();

        // Save the document to a file.
        doc.Save("TableWithStyle.docx");
    }
}
