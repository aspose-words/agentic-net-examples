using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableMargins
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Begin a new table.
            builder.StartTable();

            // Insert the first cell and add some text.
            builder.InsertCell();
            builder.Writeln("Cell with custom margins");

            // Retrieve the cell that was just created.
            Cell firstCell = builder.CurrentParagraph.ParentNode as Cell;
            if (firstCell == null)
                throw new InvalidOperationException("Unable to retrieve the created cell.");

            // Apply custom margins using cell padding properties (values are in points).
            firstCell.CellFormat.TopPadding = 10;      // Top margin
            firstCell.CellFormat.BottomPadding = 10;   // Bottom margin
            firstCell.CellFormat.LeftPadding = 15;     // Left margin
            firstCell.CellFormat.RightPadding = 15;    // Right margin

            // Insert a second cell without custom margins.
            builder.InsertCell();
            builder.Writeln("Regular cell");

            // End the current row.
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Save the document to disk.
            string outputPath = "CustomCellMargins.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output document was not created.");

            // Reload the document to confirm that margins were persisted.
            Document loadedDoc = new Document(outputPath);
            Table loadedTable = (Table)loadedDoc.GetChild(NodeType.Table, 0, true);
            Cell loadedCell = loadedTable.Rows[0].Cells[0];

            // Simple validation of one margin (padding) value.
            if (loadedCell.CellFormat.TopPadding != 10)
                throw new Exception("Cell margin (padding) was not saved correctly.");
        }
    }
}
