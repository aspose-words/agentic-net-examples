using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsRowHeightExample
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

            // Set the height of the next row to exactly 20 points.
            builder.RowFormat.Height = 20.0;
            builder.RowFormat.HeightRule = HeightRule.Exactly;

            // Second row – will have the specified height.
            builder.InsertCell();
            builder.Write("Second row, cell 1.");
            builder.InsertCell();
            builder.Write("Second row, cell 2.");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document.
            string outputPath = "RowHeight.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
