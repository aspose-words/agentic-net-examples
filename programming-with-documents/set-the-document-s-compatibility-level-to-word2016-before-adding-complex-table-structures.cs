using System;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Tables;

namespace AsposeWordsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Set the compatibility level to Word 2016.
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2016);

            // Build a complex table structure using DocumentBuilder.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start the outer table (3 rows x 3 columns).
            builder.StartTable();

            // ----- First row (simple cells) -----
            for (int i = 0; i < 3; i++)
            {
                builder.InsertCell();
                builder.Write($"Outer R1C{i + 1}");
            }
            builder.EndRow();

            // ----- Second row (merged cells + nested table) -----
            // Merge the first two cells.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Merged cells");

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // Third cell will contain a nested table.
            builder.InsertCell();
            InsertNestedTable(builder);
            builder.EndRow();

            // ----- Third row (simple cells) -----
            for (int i = 0; i < 3; i++)
            {
                builder.InsertCell();
                builder.Write($"Outer R3C{i + 1}");
            }
            builder.EndRow();

            // End the outer table.
            builder.EndTable();

            // Save the document to a file.
            doc.Save("ComplexTable.docx");
        }

        // Inserts a 2x2 nested table into the current cell.
        private static void InsertNestedTable(DocumentBuilder builder)
        {
            builder.StartTable();

            // First row of nested table.
            builder.InsertCell();
            builder.Write("Inner R1C1");
            builder.InsertCell();
            builder.Write("Inner R1C2");
            builder.EndRow();

            // Second row of nested table.
            builder.InsertCell();
            builder.Write("Inner R2C1");
            builder.InsertCell();
            builder.Write("Inner R2C2");
            builder.EndRow();

            builder.EndTable();
        }
    }
}
