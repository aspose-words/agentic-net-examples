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

        // Start a table.
        builder.StartTable();

        // Add three rows with two cells each.
        for (int rowIndex = 0; rowIndex < 3; rowIndex++)
        {
            // First cell.
            builder.InsertCell();
            builder.Writeln($"Row {rowIndex + 1}, Cell 1");
            // Second cell.
            builder.InsertCell();
            // Insert multiple lines to demonstrate automatic height expansion.
            builder.Writeln($"Row {rowIndex + 1}, Cell 2 - Line 1");
            builder.Writeln($"Row {rowIndex + 1}, Cell 2 - Line 2");
            builder.Writeln($"Row {rowIndex + 1}, Cell 2 - Line 3");

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Set each row's HeightRule to Auto so rows expand with content.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.HeightRule = HeightRule.Auto;
        }

        // Save the document.
        string outputPath = "TableAutoHeight.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
