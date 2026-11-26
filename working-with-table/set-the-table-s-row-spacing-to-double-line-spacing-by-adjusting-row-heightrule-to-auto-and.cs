using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableRowSpacingExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2x2 table.
            builder.StartTable();

            // First row, first cell.
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");

            // First row, second cell.
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            builder.EndRow();

            // Second row, first cell.
            builder.InsertCell();
            builder.Write("Row 2, Cell 1");

            // Second row, second cell.
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Retrieve the created table.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Set each row's spacing to double line spacing.
            // Assuming a default line height of 12 points, double spacing is 24 points.
            foreach (Row row in table.Rows)
            {
                row.RowFormat.HeightRule = HeightRule.Auto; // Use automatic height calculation.
                row.RowFormat.Height = 24; // Desired double line spacing in points.
            }

            // Save the document to a file.
            string outputPath = "TableRowSpacingOutput.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
