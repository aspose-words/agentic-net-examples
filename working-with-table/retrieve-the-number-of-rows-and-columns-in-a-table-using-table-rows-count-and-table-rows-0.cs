using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableInfo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a 3x4 table.
            Table table = builder.StartTable();

            for (int row = 1; row <= 3; row++)
            {
                for (int col = 1; col <= 4; col++)
                {
                    builder.InsertCell();
                    builder.Write($"R{row}C{col}");
                }
                builder.EndRow();
            }

            builder.EndTable();

            // Save the document (required by the workflow rules).
            const string outputPath = "TableInfo.docx";
            doc.Save(outputPath);

            // Retrieve the number of rows and columns.
            int rowCount = table.Rows.Count;
            int columnCount = table.Rows.Count > 0 ? table.Rows[0].Cells.Count : 0;

            // Output the results.
            Console.WriteLine($"Table has {rowCount} rows and {columnCount} columns.");
        }
    }
}
