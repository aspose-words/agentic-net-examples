using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableMergeDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a 3‑row, 2‑column table.
            builder.StartTable();

            // Row 1
            builder.InsertCell();
            builder.Write("Row 1, Col 1");
            builder.InsertCell();
            builder.Write("Row 1, Col 2");
            builder.EndRow();

            // Row 2
            builder.InsertCell();
            builder.Write("Row 2, Col 1");
            builder.InsertCell();
            builder.Write("Row 2, Col 2");
            builder.EndRow();

            // Row 3
            builder.InsertCell();
            builder.Write("Row 3, Col 1");
            builder.InsertCell();
            builder.Write("Row 3, Col 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Retrieve the created table.
            Table table = doc.FirstSection.Body.Tables[0];

            // Merge the cells of the first column vertically.
            // Set the first cell as the start of the merged range.
            Cell firstCell = table.Rows[0].Cells[0];
            firstCell.CellFormat.VerticalMerge = CellMerge.First;

            // All subsequent cells in the same column become "Previous".
            for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++)
            {
                Cell cell = table.Rows[rowIndex].Cells[0];
                cell.CellFormat.VerticalMerge = CellMerge.Previous;
            }

            // Define the output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedCells.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved successfully.");
        }
    }
}
