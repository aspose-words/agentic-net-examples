using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeCellMarginsExample
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

            // Insert the first cell.
            Cell cell = builder.InsertCell();

            // Define cell margins using padding properties (TopPadding, BottomPadding, LeftPadding, RightPadding).
            // These properties control the space between the cell borders and its contents.
            cell.CellFormat.TopPadding = 10.0;    // MarginTop
            cell.CellFormat.BottomPadding = 15.0; // MarginBottom
            cell.CellFormat.LeftPadding = 20.0;   // MarginLeft
            cell.CellFormat.RightPadding = 25.0;  // MarginRight

            // Add some text to the cell.
            builder.Write("Cell with custom margins.");

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document.
            string outputPath = "CellMargins.docx";
            doc.Save(outputPath);

            // Reload the document to verify the margins were applied.
            Document loadedDoc = new Document(outputPath);
            Table loadedTable = loadedDoc.FirstSection.Body.Tables[0];
            Cell loadedCell = loadedTable.FirstRow.FirstCell;
            CellFormat fmt = loadedCell.CellFormat;

            // Simple validation – throw if any margin does not match the expected value.
            if (fmt.TopPadding != 10.0 ||
                fmt.BottomPadding != 15.0 ||
                fmt.LeftPadding != 20.0 ||
                fmt.RightPadding != 25.0)
            {
                throw new InvalidOperationException("Cell margins were not set correctly.");
            }

            // Indicate successful completion.
            Console.WriteLine("Document created and cell margins verified successfully.");
        }
    }
}
