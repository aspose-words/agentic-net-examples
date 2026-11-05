using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with five rows.
        builder.StartTable();

        // Row 0 – normal height.
        builder.InsertCell();
        builder.Write("Row 0, Cell 0");
        builder.EndRow();

        // Row 1 – will be set to zero height (empty row).
        builder.InsertCell();
        builder.Write("Row 1, Cell 0");
        builder.EndRow();

        // Row 2 – normal height.
        builder.InsertCell();
        builder.Write("Row 2, Cell 0");
        builder.EndRow();

        // Row 3 – will be set to zero height (empty row).
        builder.InsertCell();
        builder.Write("Row 3, Cell 0");
        builder.EndRow();

        // Row 4 – normal height.
        builder.InsertCell();
        builder.Write("Row 4, Cell 0");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Set zero height for rows 1 and 3.
        table.Rows[1].RowFormat.Height = 0;
        table.Rows[1].RowFormat.HeightRule = HeightRule.Exactly;

        table.Rows[3].RowFormat.Height = 0;
        table.Rows[3].RowFormat.HeightRule = HeightRule.Exactly;

        // Remove rows whose height is zero.
        for (int i = table.Rows.Count - 1; i >= 0; i--)
        {
            Row row = table.Rows[i];
            if (row.RowFormat.Height == 0 && row.RowFormat.HeightRule == HeightRule.Exactly)
            {
                row.Remove();
            }
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
