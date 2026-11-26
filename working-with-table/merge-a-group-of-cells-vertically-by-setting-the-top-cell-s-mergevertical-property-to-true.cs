using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace MergeVerticalCellsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            builder.StartTable();

            // First row, first cell (top cell of the vertical merge).
            builder.InsertCell();
            builder.Writeln("Top Cell (merged vertically)");
            // Set the vertical merge property to First for the top cell.
            ((Cell)builder.CurrentParagraph.ParentNode).CellFormat.VerticalMerge = CellMerge.First;

            // First row, second cell.
            builder.InsertCell();
            builder.Writeln("Cell 1,2");
            builder.EndRow();

            // Second row, first cell (merged with the top cell).
            builder.InsertCell();
            builder.Writeln("Merged Cell Part 2");
            // Set the vertical merge property to Previous for the subsequent cell.
            ((Cell)builder.CurrentParagraph.ParentNode).CellFormat.VerticalMerge = CellMerge.Previous;

            // Second row, second cell.
            builder.InsertCell();
            builder.Writeln("Cell 2,2");
            builder.EndRow();

            // Third row, first cell (continue merging).
            builder.InsertCell();
            builder.Writeln("Merged Cell Part 3");
            ((Cell)builder.CurrentParagraph.ParentNode).CellFormat.VerticalMerge = CellMerge.Previous;

            // Third row, second cell.
            builder.InsertCell();
            builder.Writeln("Cell 3,2");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Save the document.
            string outputPath = "MergedCells.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
