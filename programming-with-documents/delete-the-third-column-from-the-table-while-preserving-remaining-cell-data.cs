using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace DeleteThirdColumnExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a sample table with 3 columns and 2 rows.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("R1C1");
            builder.InsertCell();
            builder.Write("R1C2");
            builder.InsertCell();
            builder.Write("R1C3");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("R2C1");
            builder.InsertCell();
            builder.Write("R2C2");
            builder.InsertCell();
            builder.Write("R2C3");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Delete the third column (index 2) from the table.
            foreach (Row row in table.Rows)
            {
                // Ensure the row has at least three cells before attempting removal.
                if (row.Cells.Count > 2)
                {
                    row.Cells.RemoveAt(2);
                }
            }

            // Save the modified document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DeletedThirdColumn.docx");
            doc.Save(outputPath);
        }
    }
}
