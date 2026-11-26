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

        // Build a simple table with two rows.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Set the first row's height rule to AtLeast with a minimum height of 15 points.
        Row firstRow = table.Rows[0];
        firstRow.RowFormat.HeightRule = HeightRule.AtLeast;
        firstRow.RowFormat.Height = 15; // Height in points.

        // Save the document to disk.
        string outputPath = "TableRowHeight.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
