using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableMergeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building a table.
            builder.StartTable();

            // ---- First row: one cell spanning the full width ----
            // Insert the first cell and mark it as the start of a horizontal merge.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Writeln("This cell spans the full width of the table.");

            // Insert additional cells in the same row and mark them as merged with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // End the first row.
            builder.EndRow();

            // ---- Optional second row: normal cells for demonstration ----
            builder.InsertCell();
            builder.Writeln("Cell 1");

            builder.InsertCell();
            builder.Writeln("Cell 2");

            builder.InsertCell();
            builder.Writeln("Cell 3");

            // End the second row.
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document to a file.
            string outputPath = "MergedTable.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output file was not created.");

            // The program ends here without waiting for user input.
        }
    }
}
