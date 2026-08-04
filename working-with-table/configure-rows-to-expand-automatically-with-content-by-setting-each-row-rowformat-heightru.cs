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

        // First row.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1.");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2.");
        // Ensure the row expands automatically.
        builder.RowFormat.HeightRule = HeightRule.Auto;
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1 with a longer text that should cause the row to expand automatically.");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2.");
        builder.RowFormat.HeightRule = HeightRule.Auto;
        builder.EndRow();

        // Third row.
        builder.InsertCell();
        builder.Write("Row 3, Cell 1.");
        builder.InsertCell();
        builder.Write("Row 3, Cell 2 with even longer content to demonstrate automatic row height adjustment.");
        builder.RowFormat.HeightRule = HeightRule.Auto;
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // As an extra safety measure, iterate all rows and set HeightRule to Auto.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.HeightRule = HeightRule.Auto;
        }

        // Save the document to a file in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RowsAutoHeight.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The document was not saved correctly.");
        }
    }
}
