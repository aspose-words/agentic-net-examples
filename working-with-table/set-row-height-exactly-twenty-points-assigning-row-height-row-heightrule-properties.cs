using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // First row (e.g., a header row).
        builder.InsertCell();
        builder.Write("Header");
        builder.EndRow();

        // Second row – we will set its height to exactly 20 points.
        builder.InsertCell();
        builder.Write("Row with fixed height");

        // Apply the height settings to the current row.
        builder.RowFormat.Height = 20;                 // Height in points.
        builder.RowFormat.HeightRule = HeightRule.Exactly; // Exact height rule.

        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        doc.Save("RowHeightExactly20.docx");
    }
}
