using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin building a table.
        Table table = builder.StartTable();

        // ----- First row (header) -----
        // Set explicit width for the first column (100 points).
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.InsertCell();
        builder.Writeln("Header 1");

        // Set explicit width for the second column (150 points).
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // ----- Second row (data) -----
        // Apply the same column widths to keep the layout fixed.
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.InsertCell();
        builder.Writeln("Row 1, Col 1");

        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
        builder.InsertCell();
        builder.Writeln("Row 1, Col 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Disable AutoFit and enforce the fixed column widths defined above.
        table.AutoFit(AutoFitBehavior.FixedColumnWidths);

        // Save the document to disk.
        doc.Save("FixedLayoutTable.docx");
    }
}
