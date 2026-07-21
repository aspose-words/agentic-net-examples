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

        // Start a new table.
        Table table = builder.StartTable();

        // Insert first cell and add some text.
        builder.InsertCell();
        builder.Write("Cell 1");

        // Insert second cell and add some text.
        builder.InsertCell();
        builder.Write("Cell 2");

        // End the first row.
        builder.EndRow();

        // Insert third cell (new row) and add text.
        builder.InsertCell();
        builder.Write("Cell 3");

        // Insert fourth cell and add text.
        builder.InsertCell();
        builder.Write("Cell 4");

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Set a fixed preferred width for the table (e.g., 300 points).
        table.PreferredWidth = PreferredWidth.FromPoints(300);

        // Center the table on the page.
        table.Alignment = TableAlignment.Center;

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FixedWidthCenteredTable.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The document was not saved correctly.");
        }
    }
}
