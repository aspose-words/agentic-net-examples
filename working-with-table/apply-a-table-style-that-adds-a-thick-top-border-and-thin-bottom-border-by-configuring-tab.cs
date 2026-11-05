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

        // Build a simple 1x1 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Sample cell");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply a thick top border.
        table.SetBorder(BorderType.Top, LineStyle.Single, 2.0, Color.Black, true);

        // Apply a thin bottom border.
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 0.5, Color.Black, true);

        // Save the document to disk.
        const string outputPath = "TableBorderExample.docx";
        doc.Save(outputPath);
    }
}
