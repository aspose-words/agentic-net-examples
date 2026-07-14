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

        // Build a 3x3 table with sample text.
        Table table = builder.StartTable();
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Delete the third column (zero‑based index 2) from each row.
        foreach (Row row in table.Rows)
        {
            if (row.Cells.Count > 2)
                row.Cells.RemoveAt(2);
        }

        // Save the modified document.
        doc.Save("DeletedThirdColumn.docx");
    }
}
