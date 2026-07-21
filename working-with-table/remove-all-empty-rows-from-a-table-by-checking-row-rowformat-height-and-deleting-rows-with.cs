using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a table with four rows. Some rows will have zero height to simulate empty rows.
            builder.StartTable();

            // Row 0 – non‑empty (height 50 points).
            builder.InsertCell();
            builder.Write("Row 0, Cell 0");
            builder.InsertCell();
            builder.Write("Row 0, Cell 1");
            builder.RowFormat.Height = 50;
            builder.RowFormat.HeightRule = HeightRule.Exactly;
            builder.EndRow();

            // Row 1 – empty (height 0 points).
            builder.InsertCell();
            builder.Write("Row 1, Cell 0");
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.RowFormat.Height = 0;
            builder.RowFormat.HeightRule = HeightRule.Exactly;
            builder.EndRow();

            // Row 2 – non‑empty (height 30 points).
            builder.InsertCell();
            builder.Write("Row 2, Cell 0");
            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.RowFormat.Height = 30;
            builder.RowFormat.HeightRule = HeightRule.Exactly;
            builder.EndRow();

            // Row 3 – empty (height 0 points).
            builder.InsertCell();
            builder.Write("Row 3, Cell 0");
            builder.InsertCell();
            builder.Write("Row 3, Cell 1");
            builder.RowFormat.Height = 0;
            builder.RowFormat.HeightRule = HeightRule.Exactly;
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the original document (optional, for inspection).
            string originalPath = Path.Combine(Environment.CurrentDirectory, "Original.docx");
            doc.Save(originalPath);

            // Retrieve the first table in the document.
            Table table = doc.FirstSection.Body.Tables[0];

            // Iterate backwards through the rows and delete those with zero height.
            for (int i = table.Rows.Count - 1; i >= 0; i--)
            {
                Row row = table.Rows[i];
                if (row.RowFormat.Height == 0)
                {
                    // Delete the row using DocumentBuilder.DeleteRow (table index = 0).
                    builder.DeleteRow(0, i);
                }
            }

            // Save the modified document.
            string resultPath = Path.Combine(Environment.CurrentDirectory, "Result.docx");
            doc.Save(resultPath);

            // Simple validation – ensure the result file was created.
            if (!File.Exists(resultPath))
                throw new InvalidOperationException("Result document was not saved correctly.");
        }
    }
}
