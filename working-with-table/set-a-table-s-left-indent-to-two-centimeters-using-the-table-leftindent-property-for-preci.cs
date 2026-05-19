using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // Insert the first cell (creates the first row) so that the table is no longer empty.
        builder.InsertCell();

        // Set the left indent to 2 cm (1 cm ≈ 28.3464567 pt).
        double pointsPerCentimeter = 72.0 / 2.54;
        table.LeftIndent = 2 * pointsPerCentimeter;

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Cell 2");

        // Finish the first row.
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Simple validation to ensure the indent was applied.
        if (Math.Abs(table.LeftIndent - 2 * pointsPerCentimeter) > 0.01)
            throw new InvalidOperationException("Failed to set the table left indent.");

        // Save the document.
        doc.Save("TableLeftIndent.docx");
    }
}
