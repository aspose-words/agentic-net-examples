using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableBottomMarginExample
{
    class Program
    {
        static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and add a simple 2x2 grid.
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
            builder.EndTable(); // Ends the table and returns the Table node.

            // Set the distance between the bottom of the table and surrounding text (bottom margin) to 5 points.
            table.DistanceBottom = 5.0;

            // Define an output path relative to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableBottomMargin.docx");

            // Save the document.
            doc.Save(outputPath);
        }
    }
}
