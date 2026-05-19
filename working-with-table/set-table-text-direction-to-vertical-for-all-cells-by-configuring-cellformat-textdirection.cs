using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableTextDirectionExample
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

            // Set the cell orientation globally before inserting cells.
            // This will affect all cells created after this point.
            builder.CellFormat.Orientation = TextOrientation.Upward;

            // Build a simple 2x2 table.
            builder.InsertCell();
            builder.Write("Cell 1,1");
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

            // Ensure that every existing cell also has vertical orientation.
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.Orientation = TextOrientation.Upward;
                }
            }

            // Define output path.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "TableVerticalText.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
