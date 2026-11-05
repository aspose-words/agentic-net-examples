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

        // Build a table with three columns.
        builder.StartTable();

        // First row – header cells.
        builder.InsertCell();
        builder.Writeln("Column 1");
        builder.InsertCell();
        builder.Writeln("Column 2");
        builder.InsertCell();
        builder.Writeln("Column 3");
        builder.EndRow();

        // Second row – data cells.
        builder.InsertCell();
        builder.Writeln("Data 1");
        builder.InsertCell();
        builder.Writeln("Data 2");
        builder.InsertCell();
        builder.Writeln("Data 3");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Define explicit column widths (in points).
        double[] columnWidths = { 100.0, 150.0, 200.0 };

        // Apply the widths to the cells of the first row – this defines column widths.
        Row firstRow = table.FirstRow;
        for (int i = 0; i < firstRow.Cells.Count && i < columnWidths.Length; i++)
        {
            Cell cell = firstRow.Cells[i];
            cell.CellFormat.Width = columnWidths[i];
        }

        // Disable AutoFit to enforce a fixed layout.
        table.AllowAutoFit = false;

        // Save the document.
        string outputPath = "FixedLayoutTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Inform that the process completed.
        Console.WriteLine("Document saved to " + Path.GetFullPath(outputPath));
    }
}
