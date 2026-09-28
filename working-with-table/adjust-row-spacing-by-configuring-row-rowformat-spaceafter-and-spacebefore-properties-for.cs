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

        // Build a simple 3‑row, 2‑column table.
        builder.StartTable();

        // Row 1
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 1");
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 2");
        builder.EndRow();

        // Row 2
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 1");
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 2");
        builder.EndRow();

        // Row 3
        builder.InsertCell();
        builder.Writeln("Row 3, Cell 1");
        builder.InsertCell();
        builder.Writeln("Row 3, Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Adjust spacing for each row.
        // The current Aspose.Words version does not expose RowFormat.SpaceBefore/SpaceAfter.
        // As an alternative, we can control the row height to simulate spacing.
        for (int i = 0; i < table.Rows.Count; i++)
        {
            Row row = table.Rows[i];

            // First row gets a larger height to simulate extra space before it.
            if (i == 0)
            {
                row.RowFormat.Height = 24.0; // points
            }
            else
            {
                row.RowFormat.Height = 12.0; // points
            }

            // Ensure the height is treated as an exact value.
            row.RowFormat.HeightRule = HeightRule.Exactly;
        }

        // Save the document.
        string outputPath = "RowSpacing.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved successfully.");
    }
}
