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

        // Build a table with two cells in a single row.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Merged Cell");
        builder.InsertCell();
        builder.Writeln("Second Cell");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Row row = table.FirstRow;
        Cell firstCell = row.Cells[0];
        Cell secondCell = row.Cells[1];

        // Merge the two adjacent cells horizontally.
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Save the document.
        string fileName = "MergedTable.docx";
        doc.Save(fileName);

        // Verify that the file was created.
        if (!File.Exists(fileName))
            throw new Exception("Failed to create the output document.");

        Console.WriteLine($"Document '{fileName}' created with merged cells.");
    }
}
