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
        Table table = builder.StartTable();

        // First row
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndTable(); // Ends the table and returns the Table node.

        // Remove any existing borders (including inner cell borders).
        table.ClearBorders();

        // Apply custom outer borders. The last parameter (true) removes any explicit cell borders,
        // ensuring that only the outer borders remain.
        table.SetBorder(BorderType.Left,   LineStyle.Single, 2.0, Color.Blue, true);
        table.SetBorder(BorderType.Right,  LineStyle.Single, 2.0, Color.Blue, true);
        table.SetBorder(BorderType.Top,    LineStyle.Single, 2.0, Color.Blue, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 2.0, Color.Blue, true);

        // Save the document.
        string outputPath = "CustomOuterBorderTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
