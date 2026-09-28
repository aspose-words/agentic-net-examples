using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to construct content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3x3 table.
        builder.StartTable();
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the created table from the document.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Apply a custom outer border (single line, 2 points, black).
        table.SetBorder(BorderType.Left, LineStyle.Single, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Right, LineStyle.Single, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Top, LineStyle.Single, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 2.0, Color.Black, true);

        // Disable inner borders (horizontal and vertical).
        table.SetBorder(BorderType.Horizontal, LineStyle.None, 0, Color.Empty, false);
        table.SetBorder(BorderType.Vertical, LineStyle.None, 0, Color.Empty, false);

        // Save the document to a file.
        string outputPath = "TableOuterBorder.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output file was not created.");
    }
}
