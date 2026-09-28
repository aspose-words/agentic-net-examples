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

        // Build a table with four rows: two with content, two empty.
        builder.StartTable();

        // Row 1 – contains text.
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 1");
        builder.EndRow();

        // Row 2 – empty row (will be marked with zero height).
        builder.InsertCell();
        // No text written – cell remains empty.
        builder.EndRow();

        // Row 3 – contains text.
        builder.InsertCell();
        builder.Writeln("Row 3, Cell 1");
        builder.EndRow();

        // Row 4 – empty row (will be marked with zero height).
        builder.InsertCell();
        // No text written.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Mark empty rows with zero height.
        foreach (Row row in table.Rows)
        {
            // A row is considered empty if its first cell contains only the end-of-cell marker.
            string cellText = row.Cells[0].GetText().Replace("\a", string.Empty).Trim();
            if (string.IsNullOrEmpty(cellText))
            {
                row.RowFormat.Height = 0;
                row.RowFormat.HeightRule = HeightRule.Exactly;
            }
        }

        // Remove rows that have zero height.
        for (int i = table.Rows.Count - 1; i >= 0; i--)
        {
            Row row = table.Rows[i];
            if (row.RowFormat.Height == 0)
            {
                table.Rows.RemoveAt(i);
            }
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
