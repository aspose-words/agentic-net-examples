using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with three columns.
        builder.StartTable();

        // First row (header).
        builder.InsertCell();
        builder.Writeln("Column 1");
        builder.InsertCell();
        builder.Writeln("Column 2");
        builder.InsertCell();
        builder.Writeln("Column 3");
        builder.EndRow();

        // Second row (data).
        builder.InsertCell();
        builder.Writeln("Data A");
        builder.InsertCell();
        builder.Writeln("Data B");
        builder.InsertCell();
        builder.Writeln("Data C");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Disable automatic fitting to use a fixed layout.
        table.AllowAutoFit = false;

        // Set explicit column widths (in points).
        // Column 1 width = 100 points, Column 2 = 150 points, Column 3 = 200 points.
        double[] columnWidths = { 100, 150, 200 };

        foreach (Row row in table.Rows)
        {
            for (int i = 0; i < row.Cells.Count; i++)
            {
                Cell cell = row.Cells[i];
                cell.CellFormat.PreferredWidth = PreferredWidth.FromPoints(columnWidths[i]);
            }
        }

        // Save the document.
        string outputPath = "FixedLayoutTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        // Confirmation message.
        Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}
