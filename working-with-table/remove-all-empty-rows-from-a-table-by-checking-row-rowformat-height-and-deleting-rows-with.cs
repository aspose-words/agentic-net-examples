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

        // Build a sample table with five rows.
        builder.StartTable();

        // Row 0 – normal height.
        builder.InsertCell();
        builder.Writeln("Row 0");
        builder.EndRow();

        // Row 1 – empty row with zero height.
        builder.InsertCell();
        // Obtain the current cell and then its parent row.
        Cell cell1 = (Cell)builder.CurrentParagraph.ParentNode;
        Row row1 = cell1.ParentRow;
        row1.RowFormat.Height = 0;
        row1.RowFormat.HeightRule = HeightRule.Exactly;
        builder.EndRow();

        // Row 2 – normal height.
        builder.InsertCell();
        builder.Writeln("Row 2");
        builder.EndRow();

        // Row 3 – empty row with zero height.
        builder.InsertCell();
        Cell cell3 = (Cell)builder.CurrentParagraph.ParentNode;
        Row row3 = cell3.ParentRow;
        row3.RowFormat.Height = 0;
        row3.RowFormat.HeightRule = HeightRule.Exactly;
        builder.EndRow();

        // Row 4 – normal height.
        builder.InsertCell();
        builder.Writeln("Row 4");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Locate the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("No table found in the document.");

        // Remove rows whose height is zero.
        for (int i = table.Rows.Count - 1; i >= 0; i--)
        {
            Row row = table.Rows[i];
            if (row.RowFormat.Height == 0 && row.RowFormat.HeightRule == HeightRule.Exactly)
                row.Remove();
        }

        // Save the resulting document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
