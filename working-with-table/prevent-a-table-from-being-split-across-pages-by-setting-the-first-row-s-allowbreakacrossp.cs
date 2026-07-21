using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableAllowBreakExample
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

            // First row – this row will not be allowed to break across pages.
            builder.InsertCell();
            builder.Write("Header Row - No break across pages");
            builder.InsertCell();
            builder.Write("Header Cell 2");
            // End the first row.
            Row firstRow = builder.EndRow();

            // Disable breaking across pages for the first row.
            firstRow.RowFormat.AllowBreakAcrossPages = false;

            // Add additional rows to make the table span multiple pages.
            for (int i = 1; i <= 30; i++)
            {
                builder.InsertCell();
                builder.Write($"Row {i}, Cell 1");
                builder.InsertCell();
                builder.Write($"Row {i}, Cell 2");
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Define the output file path.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Table.AllowBreakAcrossPages.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
