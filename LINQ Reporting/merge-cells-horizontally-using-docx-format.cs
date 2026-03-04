using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCellMergeExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a DocumentBuilder which simplifies document construction.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            Table table = builder.StartTable();

            // ---- First row: two horizontally merged cells ----
            // Insert the first cell and mark it as the start of a merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Text in merged cells.");

            // Insert the second cell and merge it with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            // No text is written to this cell because it is merged.
            builder.EndRow();

            // ---- Second row: two independent cells ----
            // Reset merge setting to None for subsequent cells.
            builder.CellFormat.HorizontalMerge = CellMerge.None;

            builder.InsertCell();
            builder.Write("Unmerged cell 1.");

            builder.InsertCell();
            builder.Write("Unmerged cell 2.");

            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Save the document in DOCX format.
            string outputPath = "CellFormat.HorizontalMerge.docx";
            doc.Save(outputPath);
        }
    }
}
