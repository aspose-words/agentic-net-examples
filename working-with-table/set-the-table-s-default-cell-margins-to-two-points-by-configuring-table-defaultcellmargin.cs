using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            Table table = builder.StartTable();

            // Set the default cell margins (padding) for all cells that will be created
            // after this point. Use the builder's CellFormat so the table already has a row.
            builder.CellFormat.SetPaddings(2.0, 2.0, 2.0, 2.0);

            // Add the first row with two cells.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Add the second row with two cells.
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document.
            string outputPath = "TableDefaultCellMargin.docx";
            doc.Save(outputPath);
        }
    }
}
