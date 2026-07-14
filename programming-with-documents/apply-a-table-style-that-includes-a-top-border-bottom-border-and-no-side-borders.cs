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

        // Start a table and add a single row with two cells.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Remove any existing borders.
        table.ClearBorders();

        // Apply a top border.
        table.SetBorder(BorderType.Top, LineStyle.Single, 1.5, Color.Black, true);
        // Apply a bottom border.
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.5, Color.Black, true);
        // Left and right borders remain unset (no side borders).

        // Save the document to the local file system.
        doc.Save("TableStyleTopBottom.docx");
    }
}
