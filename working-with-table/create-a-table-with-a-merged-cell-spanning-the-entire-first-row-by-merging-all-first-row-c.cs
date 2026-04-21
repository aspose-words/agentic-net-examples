using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableMergeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            Table table = builder.StartTable();

            // ---- First row: a single cell merged across all columns ----
            // Insert the first cell and mark it as the start of a merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Merged Header Cell");

            // Insert additional cells that will be merged with the first one.
            // The number of cells determines how many columns the merge spans.
            // Here we create a total of 4 columns.
            for (int i = 0; i < 3; i++)
            {
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            }

            // Finish the first row.
            builder.EndRow();

            // ---- Second row: normal, unmerged cells ----
            for (int col = 1; col <= 4; col++)
            {
                builder.InsertCell();
                builder.Write($"Row 2, Cell {col}");
            }
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Save the document to a local file.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedTable.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
            {
                throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");
            }

            // Optionally, inform the user (no interactive prompts required).
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
    }
}
