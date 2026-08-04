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

        // Build a simple 2x1 table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("First row, first cell.");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Second row, first cell.");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Set the height rule of the first row to Auto (no explicit height).
        // This demonstrates the required operation.
        Row firstRow = table.Rows[0];
        firstRow.RowFormat.HeightRule = HeightRule.Auto;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableRowHeightAuto.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
