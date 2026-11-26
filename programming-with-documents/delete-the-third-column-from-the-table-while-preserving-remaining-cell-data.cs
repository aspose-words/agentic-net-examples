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

        // Insert a table with 3 columns and 3 rows.
        Table table = builder.StartTable();

        // First row (header)
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.InsertCell();
        builder.Write("Header 3");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Write("Row1 Col1");
        builder.InsertCell();
        builder.Write("Row1 Col2");
        builder.InsertCell();
        builder.Write("Row1 Col3");
        builder.EndRow();

        // Third row
        builder.InsertCell();
        builder.Write("Row2 Col1");
        builder.InsertCell();
        builder.Write("Row2 Col2");
        builder.InsertCell();
        builder.Write("Row2 Col3");
        builder.EndRow();

        builder.EndTable();

        // Remove the third column (zero‑based index 2) by deleting the cell at that index in each row.
        int columnIndex = 2;
        if (table != null && table.Rows.Count > 0 && table.Rows[0].Cells.Count > columnIndex)
        {
            foreach (Row row in table.Rows)
            {
                row.Cells.RemoveAt(columnIndex);
            }
        }

        // Save the modified document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
    }
}
