using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x1 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Apply a built‑in style.
        table.Style = doc.Styles["Table Grid"];

        // Hide borders by setting each cell border style to None.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.Borders[BorderType.Left].LineStyle = LineStyle.None;
                cell.CellFormat.Borders[BorderType.Top].LineStyle = LineStyle.None;
                cell.CellFormat.Borders[BorderType.Right].LineStyle = LineStyle.None;
                cell.CellFormat.Borders[BorderType.Bottom].LineStyle = LineStyle.None;
            }
        }

        // Save the document.
        string outputPath = "TableNoBorders.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
