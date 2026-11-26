using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableFixedWidthExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            builder.StartTable();

            // First row (header).
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.InsertCell();
            builder.Write("Header 3");
            builder.EndRow();

            // Second row (data).
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            builder.InsertCell();
            builder.Write("Row 1, Cell 3");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Retrieve the created table.
            Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

            // Apply a built‑in table style.
            table.Style = doc.Styles["Table Grid"];

            // Disable AutoFit and enforce fixed column widths.
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // Define preferred widths for each column (in inches).
            // 1 inch = 72 points.
            double[] columnInches = { 1.0, 2.0, 1.5 };

            for (int col = 0; col < columnInches.Length; col++)
            {
                // Convert inches to points.
                float widthPoints = (float)(columnInches[col] * 72);

                // Apply the width to every cell in the column.
                foreach (Row row in table.Rows)
                {
                    Cell cell = row.Cells[col];
                    cell.CellFormat.PreferredWidth = PreferredWidth.FromPoints(widthPoints);
                }
            }

            // Save the document.
            string outputPath = "TableFixedWidths.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
