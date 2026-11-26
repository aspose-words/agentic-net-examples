using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableFixedWidthExample
{
    public class Program
    {
        public static void Main()
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
            Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

            // Set the table width to a fixed 15 centimeters.
            // 1 centimeter = 28.3464567 points.
            double widthInPoints = 15.0 * 28.3464567;
            // Use PreferredWidth.FromPoints for a fixed width.
            table.PreferredWidth = PreferredWidth.FromPoints(widthInPoints);

            // Save the document.
            string outputPath = "TableFixedWidth.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output file was not created.");

            // Inform that the process completed.
            Console.WriteLine("Document saved successfully to " + Path.GetFullPath(outputPath));
        }
    }
}
