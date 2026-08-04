using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableMarginExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            Table table = builder.StartTable();

            // Insert the first cell – this creates the first row and makes the table non‑empty.
            builder.InsertCell();

            // Set the default cell margins (padding) for the whole table.
            // These values will be applied to every cell that is created afterwards
            // unless a specific cell overrides them.
            table.LeftPadding = 2.0;
            table.RightPadding = 2.0;
            table.TopPadding = 2.0;
            table.BottomPadding = 2.0;

            // Fill the first cell.
            builder.Write("Cell 1,1");

            // Build the rest of the 2×2 table.
            builder.InsertCell();
            builder.Write("Cell 1,2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 2,1");
            builder.InsertCell();
            builder.Write("Cell 2,2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Table.DefaultCellMargin.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not created.");

            Console.WriteLine("Document saved to: " + outputPath);
        }
    }
}
