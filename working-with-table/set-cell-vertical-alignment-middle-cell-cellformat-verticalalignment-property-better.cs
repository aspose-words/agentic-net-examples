using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCellAlignmentExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a DocumentBuilder which will be used to insert content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            Table table = builder.StartTable();

            // Insert the first cell.
            builder.InsertCell();

            // Set the vertical alignment of the current cell to the middle (center).
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

            // Optionally set a row height so the vertical alignment effect is visible.
            builder.RowFormat.Height = 100;
            builder.RowFormat.HeightRule = HeightRule.Exactly;

            // Write some text into the cell.
            builder.Write("This text is vertically centered.");

            // Insert a second cell without changing its vertical alignment (defaults to Top).
            builder.InsertCell();
            builder.Write("This text is top-aligned.");

            // End the current row.
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Save the document to disk.
            doc.Save("CellVerticalAlignment.docx");
        }
    }
}
