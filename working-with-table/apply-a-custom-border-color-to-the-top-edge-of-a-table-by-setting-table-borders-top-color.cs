using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        builder.StartTable();

        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Apply a custom color to the top border of the table.
        // Since Table.Borders is not available, set the top border on the first row.
        Row firstRow = table.FirstRow;
        firstRow.RowFormat.Borders.Top.LineStyle = LineStyle.Single;
        firstRow.RowFormat.Borders.Top.LineWidth = 2.0; // points
        firstRow.RowFormat.Borders.Top.Color = Color.Blue; // custom border color

        // Save the document.
        string outputPath = "TableBorderTopColor.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
