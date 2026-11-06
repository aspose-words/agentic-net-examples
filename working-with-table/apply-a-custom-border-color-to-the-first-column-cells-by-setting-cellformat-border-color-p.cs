using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableBorderColorExample
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

            // Build a 3x3 table with sample text.
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    builder.InsertCell();
                    builder.Writeln($"R{row + 1}C{col + 1}");
                }
                builder.EndRow();
            }

            // End the table.
            builder.EndTable();

            // Retrieve the created table.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Define the custom border color.
            Color customColor = Color.Red;

            // Apply the custom border color to every cell in the first column.
            foreach (Row tableRow in table.Rows)
            {
                Cell firstCell = tableRow.Cells[0];
                // Set all four borders of the cell to the custom color.
                firstCell.CellFormat.Borders[BorderType.Left].Color = customColor;
                firstCell.CellFormat.Borders[BorderType.Right].Color = customColor;
                firstCell.CellFormat.Borders[BorderType.Top].Color = customColor;
                firstCell.CellFormat.Borders[BorderType.Bottom].Color = customColor;
            }

            // Save the document to disk.
            string outputPath = "FirstColumnBorderColor.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");

            // Optionally, inform that the process completed.
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
    }
}
