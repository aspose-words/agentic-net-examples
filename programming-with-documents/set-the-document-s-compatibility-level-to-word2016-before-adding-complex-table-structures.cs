using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Tables;

namespace AsposeWordsCompatibilityDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Optimize the document for Microsoft Word 2016.
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2016);

            // Use DocumentBuilder to construct complex table structures.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -------------------------------------------------
            // First table: a 3x3 table with merged cells.
            // -------------------------------------------------
            builder.StartTable();

            // Row 1
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.InsertCell();
            builder.Write("Header 3");
            builder.EndRow();

            // Row 2 with a merged cell spanning two columns.
            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Merged Cell (2 cols)");
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.EndRow();

            // Row 3
            builder.InsertCell();
            builder.Write("Row 3, Cell 1");
            builder.InsertCell();
            builder.Write("Row 3, Cell 2");
            builder.InsertCell();
            builder.Write("Row 3, Cell 3");
            builder.EndRow();

            builder.EndTable();

            // Add a paragraph break between tables.
            builder.Writeln();

            // -------------------------------------------------
            // Second table: a nested table inside a cell.
            // -------------------------------------------------
            builder.StartTable(); // Outer table

            // ----- Outer Row 1 -----
            builder.InsertCell();
            builder.Write("Outer Cell 1");

            // Cell that will contain the nested table.
            builder.InsertCell();

            // ----- Nested Table -----
            builder.StartTable(); // Nested table
            // Nested Row 1
            builder.InsertCell();
            builder.Write("Nested 1,1");
            builder.InsertCell();
            builder.Write("Nested 1,2");
            builder.EndRow();

            // Nested Row 2
            builder.InsertCell();
            builder.Write("Nested 2,1");
            builder.InsertCell();
            builder.Write("Nested 2,2");
            builder.EndRow();

            builder.EndTable(); // End of nested table

            // End outer row 1
            builder.EndRow();

            // ----- Outer Row 2 -----
            builder.InsertCell();
            builder.Write("Outer Cell 2");
            builder.InsertCell();
            builder.Write("Outer Cell 3");
            builder.EndRow();

            builder.EndTable(); // End of outer table

            // Save the document.
            string outputPath = "CompatibilityWord2016.docx";
            doc.Save(outputPath);
        }
    }
}
