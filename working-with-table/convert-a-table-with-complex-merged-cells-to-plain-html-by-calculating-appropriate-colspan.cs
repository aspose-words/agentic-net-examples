using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

namespace TableToHtmlExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a table that contains both horizontal and vertical merged cells.
            // ---------------------------------------------------------------
            // Row 1
            builder.StartTable();

            // Cell 1 – start of a horizontal merge that will span two columns.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Horizontally merged cell (col 1‑2)");

            // Cell 2 – continues the horizontal merge.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            // No text needed for the merged part.

            // Cell 3 – a normal cell.
            builder.InsertCell();
            builder.Write("Normal cell (col 3)");
            builder.EndRow();

            // Row 2
            // Cell 1 – start of a vertical merge that will span three rows.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.First;
            builder.Write("Vertically merged cell (row 2‑4)");

            // Cell 2 – normal cell.
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");

            // Cell 3 – normal cell.
            builder.InsertCell();
            builder.Write("Row 2, Cell 3");
            builder.EndRow();

            // Row 3
            // Cell 1 – continues the vertical merge.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.Previous;
            // No text needed for the merged part.

            // Cell 2 – normal cell.
            builder.InsertCell();
            builder.Write("Row 3, Cell 2");

            // Cell 3 – normal cell.
            builder.InsertCell();
            builder.Write("Row 3, Cell 3");
            builder.EndRow();

            // Row 4
            // Cell 1 – continues the vertical merge.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.Previous;
            // No text needed for the merged part.

            // Cell 2 – normal cell.
            builder.InsertCell();
            builder.Write("Row 4, Cell 2");

            // Cell 3 – normal cell.
            builder.InsertCell();
            builder.Write("Row 4, Cell 3");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document as HTML. Aspose.Words automatically generates
            // appropriate colspan and rowspan attributes for merged cells.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ComplexTable.html");
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            doc.Save(outputPath, saveOptions);

            // Simple verification that the file was created.
            if (File.Exists(outputPath))
                Console.WriteLine($"HTML file successfully created at: {outputPath}");
            else
                throw new InvalidOperationException("Failed to create the HTML output file.");
        }
    }
}
