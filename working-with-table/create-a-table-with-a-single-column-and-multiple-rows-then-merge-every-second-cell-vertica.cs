using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableMergeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table with a single column.
            builder.StartTable();

            // Add multiple rows (example: 6 rows).
            int totalRows = 6;
            for (int i = 0; i < totalRows; i++)
            {
                // Insert a cell in the single column.
                builder.InsertCell();
                // Write some text to identify the row.
                builder.Writeln($"Row {i + 1}");
                // End the current row.
                builder.EndRow();
            }

            // End the table construction.
            builder.EndTable();

            // Retrieve the created table.
            Table table = doc.FirstSection.Body.Tables[0];

            // Merge every second cell vertically (rows 2‑3, 4‑5, etc.).
            // Start merging from the second row (index 1).
            for (int rowIndex = 1; rowIndex + 1 < table.Rows.Count; rowIndex += 2)
            {
                // First cell of the pair: start of vertical merge.
                Cell firstCell = table.Rows[rowIndex].Cells[0];
                firstCell.CellFormat.VerticalMerge = CellMerge.First;

                // Second cell of the pair: continuation of vertical merge.
                Cell secondCell = table.Rows[rowIndex + 1].Cells[0];
                secondCell.CellFormat.VerticalMerge = CellMerge.Previous;
            }

            // Save the document to a file.
            string outputPath = "TableMerge.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
