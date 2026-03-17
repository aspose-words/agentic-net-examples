using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsExamples
{
    class VerticalMergeExample
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            builder.StartTable();

            // ---------- Row 1 ----------
            // Insert the first cell in the first column – this will be the start of the merged range.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.First; // Mark as the first cell in a vertical merge.
            builder.Write("Merged vertically");

            // Insert a second cell in the same row (second column) – not merged.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Row 1, Col 2");

            // End the first row.
            builder.EndRow();

            // ---------- Row 2 ----------
            // Insert a cell directly below the first cell – merge it with the cell above.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.Previous; // Continue the vertical merge.
            // No text is written to merged cells except the first one.

            // Insert a regular cell in the second column.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Row 2, Col 2");

            // End the second row.
            builder.EndRow();

            // ---------- Row 3 ----------
            // Insert another cell below the merged column – continue the merge.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.Previous;

            // Insert a regular cell in the second column.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Row 3, Col 2");

            // End the third row.
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document to disk.
            doc.Save("VerticalMergedTable.docx");
        }
    }
}
