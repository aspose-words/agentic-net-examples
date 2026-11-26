using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableBorderColorExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 3‑column, 4‑row table.
            builder.StartTable();

            for (int row = 0; row < 4; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    builder.InsertCell();
                    builder.Writeln($"R{row + 1}C{col + 1}");
                }
                builder.EndRow();
            }

            builder.EndTable();

            // Retrieve the created table.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            if (table == null)
                throw new InvalidOperationException("Table was not created.");

            // Define the custom border color.
            Color customColor = Color.FromArgb(255, 0, 128, 255); // Example: opaque orange‑blue.

            // Apply the custom border color to every cell in the first column.
            foreach (Row tableRow in table.Rows)
            {
                Cell firstCell = tableRow.Cells[0];
                // Set each side of the cell border to the custom color.
                firstCell.CellFormat.Borders[BorderType.Left].Color = customColor;
                firstCell.CellFormat.Borders[BorderType.Right].Color = customColor;
                firstCell.CellFormat.Borders[BorderType.Top].Color = customColor;
                firstCell.CellFormat.Borders[BorderType.Bottom].Color = customColor;
            }

            // Save the document.
            string outputPath = "FirstColumnBorderColor.docx";
            doc.Save(outputPath);

            // Verify that the file was saved.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not created.", outputPath);
        }
    }
}
