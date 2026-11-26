using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableShadingExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2x2 table.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndRow();
            builder.EndTable();

            // Retrieve the created table.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            if (table == null)
                throw new InvalidOperationException("Table was not created.");

            // Set the background shading of each cell to light blue.
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
                }
            }

            // Define output path.
            string outputPath = "TableShading.docx";

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was saved.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output file was not created.", outputPath);
        }
    }
}
