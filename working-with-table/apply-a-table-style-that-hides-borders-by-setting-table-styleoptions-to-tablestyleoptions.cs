using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // Insert first row with two cells.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        // Insert second row with two cells.
        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Apply a built‑in table style (optional, any style works).
        table.Style = doc.Styles["Table Grid"];

        // Hide borders by clearing them on all sides.
        table.SetBorder(BorderType.Left, LineStyle.None, 0, Color.Empty, true);
        table.SetBorder(BorderType.Right, LineStyle.None, 0, Color.Empty, true);
        table.SetBorder(BorderType.Top, LineStyle.None, 0, Color.Empty, true);
        table.SetBorder(BorderType.Bottom, LineStyle.None, 0, Color.Empty, true);
        table.SetBorder(BorderType.Horizontal, LineStyle.None, 0, Color.Empty, true);
        table.SetBorder(BorderType.Vertical, LineStyle.None, 0, Color.Empty, true);

        // Save the document.
        string outputPath = "TableNoBorders.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
