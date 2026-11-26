using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            builder.StartTable();

            // Insert the first cell.
            builder.InsertCell();

            // Get the current cell that the builder is inside.
            Cell cell = builder.CurrentParagraph.ParentNode as Cell;
            if (cell == null)
                throw new InvalidOperationException("Unable to retrieve the current cell.");

            // Set vertical alignment of the cell to middle (center).
            cell.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

            // Add some text to the cell.
            builder.Writeln("This text is vertically centered.");

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Define output file path.
            string outputPath = "VerticalAlignmentTable.docx";

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not saved.", outputPath);

            // Reload the document and verify the vertical alignment.
            Document loadedDoc = new Document(outputPath);
            Table table = loadedDoc.GetChild(NodeType.Table, 0, true) as Table;
            if (table == null)
                throw new InvalidOperationException("No table found in the saved document.");

            Cell loadedCell = table.Rows[0].Cells[0];
            if (loadedCell.CellFormat.VerticalAlignment != CellVerticalAlignment.Center)
                throw new InvalidOperationException("Cell vertical alignment is not set to center as expected.");

            // Program completed successfully.
        }
    }
}
