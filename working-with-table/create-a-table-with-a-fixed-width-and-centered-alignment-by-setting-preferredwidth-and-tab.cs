using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableFixedWidthExample
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

            // First row, first cell.
            builder.InsertCell();
            builder.Write("Cell 1");

            // First row, second cell.
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Second row, first cell.
            builder.InsertCell();
            builder.Write("Cell 3");

            // Second row, second cell.
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndRow();

            // End the table construction.
            builder.EndTable();

            // Retrieve the created table.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Set a fixed width (e.g., 400 points) and center the table.
            table.PreferredWidth = PreferredWidth.FromPoints(400);
            table.Alignment = TableAlignment.Center;

            // Save the document.
            string outputPath = "TableFixedWidth.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output document was not created.");

            // Optionally, you could add further processing here.
        }
    }
}
