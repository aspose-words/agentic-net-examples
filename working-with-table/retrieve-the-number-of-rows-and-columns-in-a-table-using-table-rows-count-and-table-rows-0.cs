using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableInfo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();

            // Build a 3x4 table using DocumentBuilder.
            DocumentBuilder builder = new DocumentBuilder(doc);
            Table table = builder.StartTable();

            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    builder.InsertCell();
                    builder.Write($"R{row + 1}C{col + 1}");
                }
                builder.EndRow();
            }

            builder.EndTable();

            // Retrieve the number of rows and columns.
            int rowCount = table.Rows.Count;
            int columnCount = rowCount > 0 ? table.Rows[0].Cells.Count : 0;

            // Output the counts.
            Console.WriteLine($"Table has {rowCount} rows and {columnCount} columns.");

            // Save the document.
            doc.Save("TableInfo.docx");
        }
    }
}
