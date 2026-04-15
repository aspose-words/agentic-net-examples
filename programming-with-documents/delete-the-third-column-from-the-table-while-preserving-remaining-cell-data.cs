using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a table with 3 columns and 3 rows.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.StartTable();
        for (int r = 0; r < 3; r++)
        {
            for (int c = 0; c < 3; c++)
            {
                builder.InsertCell();
                builder.Write($"R{r + 1}C{c + 1}");
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

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
