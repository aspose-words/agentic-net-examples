using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class HeaderShadingExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // ----- Header row -----
        // First header cell.
        builder.InsertCell();
        // Apply light gray background shading to the cell.
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
        builder.Write("Header 1");

        // Second header cell.
        builder.InsertCell();
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
        builder.Write("Header 2");

        // Finish the header row.
        builder.EndRow();

        // ----- Data rows (example) -----
        // First data row.
        builder.InsertCell();
        // Reset shading for regular cells (optional, clears previous formatting).
        builder.CellFormat.Shading.ClearFormatting();
        builder.Write("Row 1, Col 1");
        builder.InsertCell();
        builder.Write("Row 1, Col 2");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Write("Row 2, Col 1");
        builder.InsertCell();
        builder.Write("Row 2, Col 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to disk.
        doc.Save("HeaderShadedTable.docx");
    }
}
