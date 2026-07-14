using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableMerge
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // First row – first cell will be the top cell of a vertically merged range.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.First; // Mark as the first merged cell.
            builder.Write("Merged vertically");

            // First row – second cell (regular, not merged).
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Unmerged cell");
            builder.EndRow();

            // Second row – first cell merges with the cell above.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.Previous; // Merge vertically with the cell above.
            // No text is written to merged cells except the first one.
            
            // Second row – second cell (regular).
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Unmerged cell");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Define output path.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "MergedTable.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");

            // Optionally, inform that the process completed (no interactive prompts required).
            Console.WriteLine("Document created successfully at: " + outputPath);
        }
    }
}
