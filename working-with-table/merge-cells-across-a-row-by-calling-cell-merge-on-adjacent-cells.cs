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

            // Start a table.
            builder.StartTable();

            // Insert the first cell and mark it as the first cell in a merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("First part of merged cell.");

            // Insert the second cell and mark it as merged with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.Write("Second part of merged cell.");

            // Insert a third, independent cell.
            builder.CellFormat.HorizontalMerge = CellMerge.None; // Reset merge flag for new cells.
            builder.InsertCell();
            builder.Write("Separate cell.");

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Define the output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedCells.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");
        }
    }
}
