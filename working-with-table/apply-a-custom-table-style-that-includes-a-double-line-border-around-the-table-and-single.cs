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

        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply double line borders to the outer edges of the table.
        table.SetBorder(BorderType.Left, LineStyle.Double, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Right, LineStyle.Double, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Top, LineStyle.Double, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Double, 2.0, Color.Black, true);

        // Apply single line borders to the inner grid of the table.
        // Use BorderType.Horizontal and BorderType.Vertical for inside borders.
        table.SetBorder(BorderType.Horizontal, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Vertical, LineStyle.Single, 1.0, Color.Black, true);

        // Save the document to disk.
        string outputPath = "CustomTableStyle.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Indicate successful completion.
        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
    }
}
