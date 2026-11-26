using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace RowHeightExactExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple one‑row, one‑cell table.
            builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Sample cell");
            builder.EndRow();
            builder.EndTable();

            // Retrieve the first (and only) row of the table.
            Table table = doc.FirstSection.Body.Tables[0];
            Row row = table.Rows[0];

            // Set the row height to exactly 20 points.
            row.RowFormat.Height = 20.0;
            row.RowFormat.HeightRule = HeightRule.Exactly; // Correct enum value

            // Save the document to disk.
            string outputPath = "RowHeightExact.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
