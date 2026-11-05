using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableAutoFitExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2‑column table.
            builder.StartTable();

            // First cell.
            builder.InsertCell();
            builder.Writeln("First column");
            // Set the desired width for the first column.
            Cell firstCell = (Cell)builder.CurrentParagraph.ParentNode;
            firstCell.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);

            // Second cell.
            builder.InsertCell();
            builder.Writeln("Second column");
            // Set the desired width for the second column.
            Cell secondCell = (Cell)builder.CurrentParagraph.ParentNode;
            secondCell.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Retrieve the created table from the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            if (table == null)
                throw new InvalidOperationException("Table was not created.");

            // Disable automatic column resizing while preserving the widths set above.
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // Save the document to disk.
            string outputPath = "TableAutoFit.docx";
            doc.Save(outputPath);

            // Verify that the file was saved.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not created.", outputPath);
        }
    }
}
