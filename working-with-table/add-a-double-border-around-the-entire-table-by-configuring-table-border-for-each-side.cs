using System;
using System.Drawing;
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

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply a double border to each side of the table.
        table.SetBorder(BorderType.Top, LineStyle.Double, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Double, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Left, LineStyle.Double, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Right, LineStyle.Double, 2.0, Color.Black, true);

        // Save the document.
        string outputPath = "TableWithDoubleBorder.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
