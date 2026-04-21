using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableRowHeightExample
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

            // First row – default height.
            builder.InsertCell();
            builder.Write("First row, cell 1.");
            builder.InsertCell();
            builder.Write("First row, cell 2.");
            builder.EndRow();

            // Set the height rule to AtLeast and the minimum height to 15 points.
            builder.RowFormat.Height = 15.0;
            builder.RowFormat.HeightRule = HeightRule.AtLeast;

            // Second row – will use the height settings defined above.
            builder.InsertCell();
            builder.Write("Second row, cell 1.");
            builder.InsertCell();
            builder.Write("Second row, cell 2.");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Define output path (in the current working directory).
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableRowHeightAtLeast.docx");

            // Save the document.
            doc.Save(outputPath);
        }
    }
}
