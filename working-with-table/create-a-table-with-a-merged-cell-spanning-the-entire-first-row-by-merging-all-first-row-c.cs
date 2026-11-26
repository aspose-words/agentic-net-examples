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

            // Start a table.
            builder.StartTable();

            // ---- First row: a single merged cell spanning the entire row ----
            // Insert the first cell and mark it as the start of a horizontal merge.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Writeln("Merged Header");

            // Insert additional cells for the same row and mark them as continuation of the merge.
            // Assuming a total of 3 columns for demonstration.
            for (int i = 1; i < 3; i++)
            {
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                // Content of merged cells can be empty.
                builder.Writeln(string.Empty);
            }

            // End the first row.
            builder.EndRow();

            // Reset the merge setting for subsequent rows.
            builder.CellFormat.HorizontalMerge = CellMerge.None;

            // ---- Second row: normal cells ----
            for (int i = 0; i < 3; i++)
            {
                builder.InsertCell();
                builder.Writeln($"Cell {i + 1}");
            }

            // End the second row.
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Save the document to a file.
            string outputPath = "MergedTable.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
