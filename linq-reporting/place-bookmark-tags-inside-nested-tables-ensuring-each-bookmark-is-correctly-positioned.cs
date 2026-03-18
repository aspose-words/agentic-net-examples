using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace BookmarkInNestedTables
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ------------------------------------------------------------
            // Build the outer table (2 rows x 2 columns).
            // ------------------------------------------------------------
            builder.StartTable();

            // First row, first cell.
            builder.InsertCell();
            // Insert a bookmark that starts before the inner table.
            builder.StartBookmark("OuterCell1_BeforeInner");
            builder.Writeln("Before inner table");
            // End the bookmark.
            builder.EndBookmark("OuterCell1_BeforeInner");

            // Insert the inner table inside this cell.
            InsertInnerTable(builder, "InnerTable1");

            // First row, second cell.
            builder.InsertCell();
            // Insert a bookmark that starts after the inner table.
            builder.StartBookmark("OuterCell2_AfterInner");
            builder.Writeln("After inner table");
            builder.EndBookmark("OuterCell2_AfterInner");

            // End first row.
            builder.EndRow();

            // ------------------------------------------------------------
            // Second row – demonstrate a bookmark that spans the whole cell.
            // ------------------------------------------------------------
            // Second row, first cell.
            builder.InsertCell();
            // Bookmark that encloses the entire cell content (including inner table).
            builder.StartBookmark("OuterCell3_FullSpan");
            builder.Writeln("Start of full-span cell");
            InsertInnerTable(builder, "InnerTable2");
            builder.Writeln("End of full-span cell");
            builder.EndBookmark("OuterCell3_FullSpan");

            // Second row, second cell – empty cell with a simple bookmark.
            builder.InsertCell();
            builder.StartBookmark("OuterCell4_Simple");
            builder.Writeln("Simple bookmark content");
            builder.EndBookmark("OuterCell4_Simple");

            // End the outer table.
            builder.EndTable();

            // ------------------------------------------------------------
            // Save the document.
            // ------------------------------------------------------------
            doc.Save("BookmarksInNestedTables.docx");
        }

        /// <summary>
        /// Inserts a 2x2 inner table into the current cell.
        /// Each cell of the inner table receives its own bookmark.
        /// </summary>
        /// <param name="builder">The DocumentBuilder positioned inside a cell.</param>
        /// <param name="innerTableId">A unique identifier used to name the bookmarks.</param>
        private static void InsertInnerTable(DocumentBuilder builder, string innerTableId)
        {
            // Start the inner table.
            builder.StartTable();

            // Row 1, Cell 1
            builder.InsertCell();
            builder.StartBookmark($"{innerTableId}_R1C1");
            builder.Writeln("Inner R1C1");
            builder.EndBookmark($"{innerTableId}_R1C1");

            // Row 1, Cell 2
            builder.InsertCell();
            builder.StartBookmark($"{innerTableId}_R1C2");
            builder.Writeln("Inner R1C2");
            builder.EndBookmark($"{innerTableId}_R1C2");

            // End first row.
            builder.EndRow();

            // Row 2, Cell 1
            builder.InsertCell();
            builder.StartBookmark($"{innerTableId}_R2C1");
            builder.Writeln("Inner R2C1");
            builder.EndBookmark($"{innerTableId}_R2C1");

            // Row 2, Cell 2
            builder.InsertCell();
            builder.StartBookmark($"{innerTableId}_R2C2");
            builder.Writeln("Inner R2C2");
            builder.EndBookmark($"{innerTableId}_R2C2");

            // End second row.
            builder.EndRow();

            // End the inner table.
            builder.EndTable();
        }
    }
}
