using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTablePerformance
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to construct the table.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and add a header row.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Index");
            builder.InsertCell();
            builder.Write("Value");
            builder.EndRow();

            // IMPORTANT: Do NOT call UpdatePageLayout inside the loop.
            // Adding many rows without triggering layout recalculation greatly improves performance.
            const int rowCount = 5000; // Simulate a large table.
            for (int i = 1; i <= rowCount; i++)
            {
                builder.InsertCell();
                builder.Write(i.ToString());

                builder.InsertCell();
                builder.Write($"Row {i} data");

                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // After all batch modifications are complete, rebuild the layout once.
            doc.UpdatePageLayout();

            // Save the document.
            doc.Save("LargeTableOptimized.docx");
        }
    }
}
