using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCellPaddingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2x2 table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Set padding of 3 points on all sides for every cell in the table.
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    // Using the SetPaddings method for brevity.
                    cell.CellFormat.SetPaddings(3, 3, 3, 3);
                }
            }

            // Define output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CellPadding.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output document was not saved correctly.");
        }
    }
}
