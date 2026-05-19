using System;
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
        builder.Write("R1C1");
        builder.InsertCell();
        builder.Write("R1C2");
        builder.EndRow();

        // Row 2
        builder.InsertCell();
        builder.Write("R2C1");
        builder.InsertCell();
        builder.Write("R2C2");
        builder.EndRow();

        // Row 3
        builder.InsertCell();
        builder.Write("R3C1");
        builder.InsertCell();
        builder.Write("R3C2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = doc.FirstSection.Body.Tables[0];

        // Adjust spacing for each row.
        // Since RowFormat does not expose SpaceBefore/SpaceAfter in the current API,
        // we simulate spacing by setting a minimum row height.
        // The Height value (in points) adds space inside the row; using HeightRule.AtLeast
        // ensures the row will be at least this tall, effectively creating visual spacing.
        foreach (Row row in table.Rows)
        {
            // Desired total spacing (e.g., 10 points before + 5 points after = 15 points).
            row.RowFormat.Height = 15;
            row.RowFormat.HeightRule = HeightRule.AtLeast;
        }

        // Save the document to the local file system.
        doc.Save("RowSpacing.docx");
    }
}
