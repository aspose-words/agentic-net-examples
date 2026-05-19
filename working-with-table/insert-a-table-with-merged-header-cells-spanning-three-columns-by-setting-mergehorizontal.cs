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

            // Start the table.
            Table table = builder.StartTable();

            // ----- Header row (merged across three columns) -----
            // First cell: start horizontal merge.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Header spanning three columns");

            // Second cell: merge with previous.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // Third cell: merge with previous.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // End the header row.
            builder.EndRow();

            // Reset merge setting for subsequent rows.
            builder.CellFormat.HorizontalMerge = CellMerge.None;

            // ----- Data row (three separate cells) -----
            builder.InsertCell();
            builder.Write("Column 1");
            builder.InsertCell();
            builder.Write("Column 2");
            builder.InsertCell();
            builder.Write("Column 3");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Prepare output folder and file path.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            string outputPath = Path.Combine(outputDir, "MergedHeaderTable.docx");

            // Save the document.
            doc.Save(outputPath);
        }
    }
}
