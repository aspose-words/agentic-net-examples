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

        // Build a 3x3 table.
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

        // Save the original table for reference.
        doc.Save("TableBefore.docx");

        // Delete the second column (index 1) by removing the cell at that index from each row.
        if (table.Rows.Count > 0 && table.Rows[0].Cells.Count > 1)
        {
            foreach (Row row in table.Rows)
            {
                // Remove the cell at column index 1.
                row.Cells.RemoveAt(1);
            }
        }

        // Save the document after column removal.
        doc.Save("TableAfter.docx");
    }
}
