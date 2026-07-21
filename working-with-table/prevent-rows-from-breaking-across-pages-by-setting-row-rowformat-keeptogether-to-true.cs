using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableRowKeepTogether
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

            // Add a header row (optional) and set it to repeat on each page.
            builder.RowFormat.HeadingFormat = true;
            builder.InsertCell();
            builder.Write("Header Column 1");
            builder.InsertCell();
            builder.Write("Header Column 2");
            builder.EndRow();

            // Disable the ability for rows to break across pages.
            // This will keep each row together on a single page.
            builder.RowFormat.AllowBreakAcrossPages = false;

            // Add enough rows with long text to force the table to span multiple pages.
            for (int i = 1; i <= 30; i++)
            {
                builder.InsertCell();
                // Repeat a long string to increase the row height.
                builder.Write($"Row {i}, Column 1. " + new string('A', 500));
                builder.InsertCell();
                builder.Write($"Row {i}, Column 2. " + new string('B', 500));
                builder.EndRow();
            }

            // End the table.
            builder.EndTable();

            // Save the document to the local file system.
            string outputPath = "TableRowKeepTogether.docx";
            doc.Save(outputPath);
        }
    }
}
