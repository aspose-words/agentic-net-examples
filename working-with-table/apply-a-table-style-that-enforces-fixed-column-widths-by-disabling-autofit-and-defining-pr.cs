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

        // Build a simple 3‑column table.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.InsertCell();
        builder.Writeln("Header 3");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("Data A");
        builder.InsertCell();
        builder.Writeln("Data B");
        builder.InsertCell();
        builder.Writeln("Data C");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table safely.
        Table table = doc.GetChild(NodeType.Table, 0, true) as Table;
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Disable AutoFit and enforce fixed column widths.
        table.AutoFit(AutoFitBehavior.FixedColumnWidths);

        // Set a preferred width of 100 points for each column.
        foreach (Cell cell in table.Rows[0].Cells)
        {
            cell.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        }

        // Save the document.
        string fileName = "FixedWidthTable.docx";
        doc.Save(fileName);

        // Verify that the file was created.
        if (!File.Exists(fileName))
            throw new Exception("The output document was not created.");

        Console.WriteLine($"Document saved successfully as '{fileName}'.");
    }
}
