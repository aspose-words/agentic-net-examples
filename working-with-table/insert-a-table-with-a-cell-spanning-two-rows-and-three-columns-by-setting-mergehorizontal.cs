using System;
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

            // Start the table.
            Table table = builder.StartTable();

            // ---------- First Row ----------
            // First cell – the top‑left cell that will span 2 rows and 3 columns.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.CellFormat.VerticalMerge = CellMerge.First;
            builder.Write("Merged Cell (2 rows x 3 columns)");

            // Second cell – part of the horizontal merge.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write(string.Empty);

            // Third cell – part of the horizontal merge.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write(string.Empty);

            // End the first row.
            builder.EndRow();

            // ---------- Second Row ----------
            // First cell – continues the vertical merge.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.CellFormat.VerticalMerge = CellMerge.Previous;
            builder.Write(string.Empty);

            // Second cell – continues both horizontal and vertical merge.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.CellFormat.VerticalMerge = CellMerge.Previous;
            builder.Write(string.Empty);

            // Third cell – continues both horizontal and vertical merge.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.CellFormat.VerticalMerge = CellMerge.Previous;
            builder.Write(string.Empty);

            // End the second row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document to the local file system.
            const string outputPath = "MergedTable.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to {outputPath}");
        }
    }
}
