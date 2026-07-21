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

        // Build a 3x3 table with sample data.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.InsertCell();
        builder.Write("Header 3");
        builder.EndRow();

        // Two data rows.
        for (int i = 1; i <= 2; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i} Col 1");
            builder.InsertCell();
            builder.Write($"Row {i} Col 2");
            builder.InsertCell();
            builder.Write($"Row {i} Col 3");
            builder.EndRow();
        }

        builder.EndTable();

        // Delete the third column (zero‑based index 2) from the table.
        foreach (Row row in table.Rows)
        {
            if (row.Cells.Count > 2)
            {
                row.Cells.RemoveAt(2);
            }
        }

        // Save the modified document.
        doc.Save("DeletedThirdColumn.docx");
    }
}
