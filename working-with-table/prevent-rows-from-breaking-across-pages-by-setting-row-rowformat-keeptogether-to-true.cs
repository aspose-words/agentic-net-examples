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

        // Build a simple table with three rows and two columns.
        builder.StartTable();

        // First row
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // Third row
        builder.InsertCell();
        builder.Write("Row 3, Cell 1");
        builder.InsertCell();
        builder.Write("Row 3, Cell 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table from the document.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Prevent each row from breaking across pages.
        foreach (Row row in table.Rows)
        {
            // Setting AllowBreakAcrossPages to false keeps the row together.
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the document.
        string outputPath = "KeepTogetherRows.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
