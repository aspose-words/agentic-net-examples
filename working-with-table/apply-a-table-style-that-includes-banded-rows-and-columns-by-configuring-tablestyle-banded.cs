using System;
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

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Create a new table style.
        TableStyle bandedStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "BandedStyle");
        // Inherit basic formatting from an existing style (optional).
        bandedStyle.BaseStyleName = "Table Grid";

        // Apply the new style to the table.
        table.StyleName = "BandedStyle";

        // Enable banded rows and columns via TableStyleOptions.
        table.StyleOptions = TableStyleOptions.RowBands | TableStyleOptions.ColumnBands;

        // Save the document.
        string outputPath = "BandedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Inform the user (non‑interactive).
        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
    }
}
