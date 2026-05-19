using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableRowHeightAuto
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

            // First row.
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            // End the first row.
            Row firstRow = builder.EndRow();

            // Set the height rule of the first row to Auto (no explicit height).
            firstRow.RowFormat.HeightRule = HeightRule.Auto;

            // Second row (for comparison, set a fixed height).
            builder.InsertCell();
            builder.RowFormat.Height = 50;               // Explicit height.
            builder.RowFormat.HeightRule = HeightRule.Exactly;
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableRowHeightAuto.docx");
            doc.Save(outputPath);
        }
    }
}
