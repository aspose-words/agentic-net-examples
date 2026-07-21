using System;
using System.IO;
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

            // Adjust row height as a simple way to add visual spacing.
            // The RowFormat class does not expose SpaceBefore/SpaceAfter properties,
            // so we use the Height property with the AtLeast rule to create extra space.
            double[] extraHeights = { 5, 10, 15 }; // points of additional space per row

            for (int i = 0; i < table.Rows.Count; i++)
            {
                Row row = table.Rows[i];
                row.RowFormat.Height = extraHeights[i];
                row.RowFormat.HeightRule = HeightRule.AtLeast;
            }

            // Save the document to the local folder.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "RowSpacing.docx");
            doc.Save(outputPath);
        }
    }
}
