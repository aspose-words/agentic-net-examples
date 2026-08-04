using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCellMarginsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // First cell with custom paddings (margins inside the cell).
            Cell cell1 = builder.InsertCell();
            builder.Write("First cell with custom margins.");
            cell1.CellFormat.TopPadding = 10;      // 10 points top padding
            cell1.CellFormat.BottomPadding = 10;   // 10 points bottom padding
            cell1.CellFormat.LeftPadding = 15;     // 15 points left padding
            cell1.CellFormat.RightPadding = 15;    // 15 points right padding

            // Second cell with different custom paddings.
            Cell cell2 = builder.InsertCell();
            builder.Write("Second cell with different margins.");
            cell2.CellFormat.TopPadding = 5;
            cell2.CellFormat.BottomPadding = 5;
            cell2.CellFormat.LeftPadding = 8;
            cell2.CellFormat.RightPadding = 8;

            // End the first row.
            builder.EndRow();

            // Add a second row to demonstrate that paddings are per‑cell.
            Cell cell3 = builder.InsertCell();
            builder.Write("Third cell (default margins).");

            Cell cell4 = builder.InsertCell();
            builder.Write("Fourth cell (default margins).");

            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "CustomCellMargins.docx");
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");
        }
    }
}
