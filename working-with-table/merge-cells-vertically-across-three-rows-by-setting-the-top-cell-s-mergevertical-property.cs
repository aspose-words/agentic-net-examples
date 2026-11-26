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

            // Start a table with three rows and two columns.
            builder.StartTable();

            // ---------- Row 1 ----------
            // First cell – this will be the top cell of the vertical merge.
            builder.InsertCell();
            builder.Write("Merged Cell");
            // Set the vertical merge type to First for the top cell.
            Cell topCell = (Cell)builder.CurrentParagraph.ParentNode;
            topCell.CellFormat.VerticalMerge = CellMerge.First;

            // Second cell – regular content.
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            builder.EndRow();

            // ---------- Row 2 ----------
            // First cell – continues the vertical merge.
            builder.InsertCell();
            Cell mergeCell2 = (Cell)builder.CurrentParagraph.ParentNode;
            mergeCell2.CellFormat.VerticalMerge = CellMerge.Previous;
            builder.Write(string.Empty); // Empty content for merged part.

            // Second cell – regular content.
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.EndRow();

            // ---------- Row 3 ----------
            // First cell – continues the vertical merge.
            builder.InsertCell();
            Cell mergeCell3 = (Cell)builder.CurrentParagraph.ParentNode;
            mergeCell3.CellFormat.VerticalMerge = CellMerge.Previous;
            builder.Write(string.Empty); // Empty content for merged part.

            // Second cell – regular content.
            builder.InsertCell();
            builder.Write("Row 3, Cell 2");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Save the document.
            string outputPath = "MergedTable.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception($"Failed to create the output file: {outputPath}");

            // Optionally, inform that the process completed.
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
    }
}
