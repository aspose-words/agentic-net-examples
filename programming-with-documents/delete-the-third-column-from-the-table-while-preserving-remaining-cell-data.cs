using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class DeleteThirdColumnExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a sample 3‑column table with two rows.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("R1C1");
        builder.InsertCell();
        builder.Write("R1C2");
        builder.InsertCell();
        builder.Write("R1C3");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("R2C1");
        builder.InsertCell();
        builder.Write("R2C2");
        builder.InsertCell();
        builder.Write("R2C3");
        builder.EndRow();

        builder.EndTable();

        // Delete the third column (index 2) from the table.
        // Iterate over each row and remove the cell at the target index.
        foreach (Row row in table.Rows)
        {
            if (row.Cells.Count > 2)
            {
                row.Cells.RemoveAt(2);
            }
        }

        // Save the modified document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DeletedColumn.docx");
        doc.Save(outputPath);
    }
}
