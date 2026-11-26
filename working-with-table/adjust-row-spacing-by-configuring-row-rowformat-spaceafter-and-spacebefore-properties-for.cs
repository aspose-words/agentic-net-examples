using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace RowSpacingExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple table with three rows and two cells per row.
            builder.StartTable();

            for (int rowIndex = 0; rowIndex < 3; rowIndex++)
            {
                // First cell of the current row.
                builder.InsertCell();
                builder.Writeln($"Row {rowIndex + 1} - Cell 1");

                // Second cell of the current row.
                builder.InsertCell();
                builder.Writeln($"Row {rowIndex + 1} - Cell 2");

                // End the current row.
                builder.EndRow();
            }

            // End the table.
            builder.EndTable();

            // Retrieve the created table from the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Adjust spacing before and after each row.
            // Since RowFormat does not expose SpaceBefore/SpaceAfter, we simulate the effect
            // by setting the top and bottom cell padding for each row.
            for (int i = 0; i < table.Rows.Count; i++)
            {
                Row row = table.Rows[i];
                float spaceBefore = i * 5f; // 0pt, 5pt, 10pt...
                float spaceAfter = (table.Rows.Count - i - 1) * 5f; // 10pt, 5pt, 0pt...

                foreach (Cell cell in row.Cells)
                {
                    // Apply padding to simulate spacing.
                    cell.CellFormat.TopPadding = spaceBefore;
                    cell.CellFormat.BottomPadding = spaceAfter;
                }
            }

            // Save the document to disk.
            string outputPath = "RowSpacing.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
