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

            // Build a simple 3x3 table.
            Table table = builder.StartTable();

            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    builder.InsertCell();
                    builder.Write($"R{row + 1}C{col + 1}");
                }
                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Ensure each row stays together on the same page.
            // The RowFormat does not have a KeepTogether property; setting AllowBreakAcrossPages to false
            // prevents the row from being split across pages, achieving the same effect.
            foreach (Row r in table.Rows)
            {
                r.RowFormat.AllowBreakAcrossPages = false;
            }

            // Define the output path.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            string outputPath = Path.Combine(outputDir, "TableKeepTogether.docx");

            // Save the document.
            doc.Save(outputPath);
        }
    }
}
