using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading.
        builder.Writeln("Table with pagination (AllowBreakAcrossPages = true)");

        // Start the table.
        builder.StartTable();

        // Create 50 rows with two cells each.
        for (int i = 0; i < 50; i++)
        {
            // First cell.
            builder.InsertCell();
            builder.Write($"Row {i + 1} - Cell 1");

            // Second cell.
            builder.InsertCell();
            builder.Write($"Row {i + 1} - Cell 2");

            // End the row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Enable breaking across pages for each row and set a fixed height.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = true;
            row.RowFormat.Height = 20; // Height in points.
            row.RowFormat.HeightRule = HeightRule.Exactly;
        }

        // Save the document.
        string outputPath = "PaginationTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // Indicate success (no interactive input required).
        Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}
