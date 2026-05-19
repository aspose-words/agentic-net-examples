using System;
using System.IO;
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

        // ----- First row -----
        // Set the row height to 24 points (approximately double line spacing)
        // and use HeightRule.Auto so the height can grow if needed.
        builder.RowFormat.Height = 24.0;
        builder.RowFormat.HeightRule = HeightRule.Auto;

        // Add two cells to the first row.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        // Finish the first row.
        builder.EndRow();

        // ----- Second row (default formatting) -----
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableRowSpacing.docx");
        doc.Save(outputPath);
    }
}
