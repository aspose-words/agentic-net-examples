using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("Cell 1,1");
        builder.InsertCell();
        builder.Writeln("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Cell 2,1");
        builder.InsertCell();
        builder.Writeln("Cell 2,2");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply a custom outer border (single line, 2 pt, black).
        table.SetBorder(BorderType.Left,   LineStyle.Single, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Top,    LineStyle.Single, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Right,  LineStyle.Single, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 2.0, Color.Black, true);

        // Disable inner (horizontal and vertical) borders.
        table.SetBorder(BorderType.Horizontal, LineStyle.None, 0, Color.Empty, false);
        table.SetBorder(BorderType.Vertical,   LineStyle.None, 0, Color.Empty, false);

        // Save the document.
        string outputPath = "TableOuterBorder.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
