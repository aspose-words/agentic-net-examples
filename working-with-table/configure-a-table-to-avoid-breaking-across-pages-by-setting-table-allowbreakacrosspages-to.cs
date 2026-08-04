using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building a table.
            Table table = builder.StartTable();

            // Add multiple rows with two cells each.
            for (int i = 0; i < 20; i++)
            {
                builder.InsertCell();
                builder.Write($"Row {i + 1}, Cell 1");

                builder.InsertCell();
                builder.Write($"Row {i + 1}, Cell 2");

                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Configure each row so it will not break across pages.
            foreach (Row row in table.Rows)
            {
                row.RowFormat.AllowBreakAcrossPages = false;
            }

            // Save the document to the local file system.
            doc.Save("Table_NoBreakAcrossPages.docx");
        }
    }
}
