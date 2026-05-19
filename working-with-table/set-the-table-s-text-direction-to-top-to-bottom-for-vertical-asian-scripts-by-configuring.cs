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

            // Build a 2x2 table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1, Row 1");
            builder.InsertCell();
            builder.Write("Cell 2, Row 1");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 1, Row 2");
            builder.InsertCell();
            builder.Write("Cell 2, Row 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Set text direction for each cell to emulate top‑to‑bottom orientation
            // for vertical Asian scripts.
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.Orientation = TextOrientation.VerticalFarEast;
                }
            }

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableTextDirection.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
