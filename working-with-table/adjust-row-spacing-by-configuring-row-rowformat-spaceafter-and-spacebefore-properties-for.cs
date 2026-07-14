using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace RowSpacingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 3‑row, 2‑column table.
            Table table = builder.StartTable();

            // Row 1
            builder.InsertCell();
            builder.Write("R1C1");
            builder.InsertCell();
            builder.Write("R1C2");
            builder.EndRow();

            // Row 2
            builder.InsertCell();
            builder.Write("R2C1");
            builder.InsertCell();
            builder.Write("R2C2");
            builder.EndRow();

            // Row 3
            builder.InsertCell();
            builder.Write("R3C1");
            builder.InsertCell();
            builder.Write("R3C2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Adjust the height of each row to simulate spacing.
            // Height is measured in points. Here we increase the height progressively.
            for (int i = 0; i < table.Rows.Count; i++)
            {
                Row row = table.Rows[i];
                row.RowFormat.Height = (i + 1) * 5.0;          // 5, 10, 15 points
                row.RowFormat.HeightRule = HeightRule.AtLeast; // Ensure the row can grow if needed
            }

            // Save the document to the local file system.
            string outputPath = "RowSpacing.docx";
            doc.Save(outputPath);
        }
    }
}
