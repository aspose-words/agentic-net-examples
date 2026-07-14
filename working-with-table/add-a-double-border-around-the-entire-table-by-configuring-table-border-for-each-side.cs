using System;
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

        // Start building a simple 2x2 table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply a double line border to each side of the table.
        table.SetBorder(BorderType.Left, LineStyle.Double, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Right, LineStyle.Double, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Top, LineStyle.Double, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Double, 2.0, Color.Black, true);

        // Save the document to the local file system.
        doc.Save("Table.DoubleBorder.docx");
    }
}
