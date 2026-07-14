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

            // Start a new table.
            Table table = builder.StartTable();

            // -------------------------------------------------
            // First row – simple three separate cells (header).
            // -------------------------------------------------
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.InsertCell();
            builder.Write("Header 3");
            builder.EndRow();

            // -------------------------------------------------
            // Second row – cells with varying horizontal spans.
            // -------------------------------------------------

            // Cell that spans two columns.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First; // start of merge range
            builder.Write("Span 2 columns");
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous; // merged with previous cell
            // No text needed for merged cell.

            // Reset merge setting for the next independent cell.
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.InsertCell();
            builder.Write("Normal cell");

            // Cell that spans three columns.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Span 3 columns");
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            // End of merged range – no text for the merged cells.

            // Finish the row.
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Save the document to a local file.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "MergedCells.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");

            // Optionally, inform that the process completed (no user input required).
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
    }
}
