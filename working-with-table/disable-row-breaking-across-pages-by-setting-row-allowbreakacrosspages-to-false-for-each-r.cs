using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableRowBreakExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a sample table with several rows.
            Table table = builder.StartTable();

            // Add 5 rows, each with a long text to increase the chance of page breaks.
            for (int i = 1; i <= 5; i++)
            {
                builder.InsertCell();
                builder.Write($"Row {i} - This is a long piece of text that is intended to span multiple lines and potentially cause the row to break across pages if allowed. " +
                              "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Disable row breaking across pages for every row in the table.
            foreach (Row row in table.Rows)
            {
                row.RowFormat.AllowBreakAcrossPages = false;
            }

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Table.AllowBreakAcrossPages.docx");
            doc.Save(outputPath);
        }
    }
}
