using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

namespace AsposeWordsTableMergeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building a table.
            Table table = builder.StartTable();

            // ---------- First Row ----------
            // Insert first cell – this will be the start of a vertical merge.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.First;   // Mark as first cell in vertical merge.
            builder.CellFormat.HorizontalMerge = CellMerge.First; // Also start a horizontal merge.
            builder.Write("Vertically & Horizontally Merged");

            // Insert second cell – merge horizontally with the previous cell.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;   // No vertical merge for this cell.
            builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Merge horizontally.
            // No text needed for merged cell.

            // Insert third cell – normal, unmerged.
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.InsertCell();
            builder.Write("Normal Cell");

            // End the first row.
            builder.EndRow();

            // ---------- Second Row ----------
            // Insert first cell – merge vertically with the cell above.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.Previous; // Continue vertical merge.
            builder.CellFormat.HorizontalMerge = CellMerge.None;   // No horizontal merge here.
            // No text needed for merged cell.

            // Insert second cell – normal cell.
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.InsertCell();
            builder.Write("Second Row, Cell 2");

            // Insert third cell – normal cell.
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.InsertCell();
            builder.Write("Second Row, Cell 3");

            // End the second row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document as a Markdown file.
            // The table will be exported as Markdown because it does not contain complex structures.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                // Export tables as Markdown (default). No need to set ExportAsHtml.
                TableContentAlignment = TableContentAlignment.Auto
            };

            doc.Save("MergedTable.md", saveOptions);
        }
    }
}
