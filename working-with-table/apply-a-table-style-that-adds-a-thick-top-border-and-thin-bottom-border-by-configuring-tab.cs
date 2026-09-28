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
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Apply a thick top border to the first row.
        Row firstRow = table.FirstRow;
        firstRow.RowFormat.Borders.Top.LineWidth = 2.0;    // Thick border (points)
        firstRow.RowFormat.Borders.Top.LineStyle = LineStyle.Single;

        // Apply a thin bottom border to the last row.
        Row lastRow = table.LastRow;
        lastRow.RowFormat.Borders.Bottom.LineWidth = 0.5; // Thin border (points)
        lastRow.RowFormat.Borders.Bottom.LineStyle = LineStyle.Single;

        // Save the document.
        string outputPath = "TableBorders.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
