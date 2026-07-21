using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table with a single column.
            Table table = builder.StartTable();

            int totalRows = 6; // Number of rows in the table.

            for (int rowIndex = 1; rowIndex <= totalRows; rowIndex++)
            {
                // Insert the only cell for this row.
                builder.InsertCell();

                // Determine vertical merge settings.
                if (rowIndex == 1 || rowIndex == 6)
                {
                    // First and last rows are not merged.
                    builder.CellFormat.VerticalMerge = CellMerge.None;
                    builder.Write($"Row {rowIndex}");
                }
                else if (rowIndex % 2 == 0)
                {
                    // Even rows start a merged group.
                    builder.CellFormat.VerticalMerge = CellMerge.First;
                    builder.Write($"Group {(rowIndex / 2)}");
                }
                else
                {
                    // Odd rows (after an even row) continue the merge.
                    builder.CellFormat.VerticalMerge = CellMerge.Previous;
                    // No text is written to merged cells other than the first.
                }

                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableVerticalMerge.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");
        }
    }
}
