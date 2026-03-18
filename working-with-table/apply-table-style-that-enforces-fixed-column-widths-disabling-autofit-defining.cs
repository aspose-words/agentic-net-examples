using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document and a DocumentBuilder to construct its contents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a header row with three columns.
        Table table = builder.StartTable();

        builder.InsertCell();
        builder.Write("Column 1");
        builder.InsertCell();
        builder.Write("Column 2");
        builder.InsertCell();
        builder.Write("Column 3");
        builder.EndRow();

        // Add a second row with sample data.
        builder.InsertCell();
        builder.Write("Data 1");
        builder.InsertCell();
        builder.Write("Data 2");
        builder.InsertCell();
        builder.Write("Data 3");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Disable AutoFit so the table uses fixed column widths.
        table.AutoFit(AutoFitBehavior.FixedColumnWidths);

        // Set explicit widths for each column (in points). Applying the width to the first row's cells
        // propagates the setting to the entire column.
        table.FirstRow.Cells[0].CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        table.FirstRow.Cells[1].CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
        table.FirstRow.Cells[2].CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);

        // Save the document.
        doc.Save("FixedColumnWidthsTable.docx");
    }
}
