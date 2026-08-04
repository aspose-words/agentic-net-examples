using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace PreventRowBreakAcrossPagesExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple table with three rows.
            Table table = builder.StartTable();

            // Row 1
            builder.InsertCell();
            builder.Write("This is a long piece of text in the first cell. " +
                          "It is intended to be long enough to potentially span multiple lines.");
            builder.InsertCell();
            builder.Write("Second cell, first row.");
            builder.EndRow();

            // Row 2
            builder.InsertCell();
            builder.Write("Second row, first cell with more text to illustrate the setting.");
            builder.InsertCell();
            builder.Write("Second row, second cell.");
            builder.EndRow();

            // Row 3
            builder.InsertCell();
            builder.Write("Third row, first cell.");
            builder.InsertCell();
            builder.Write("Third row, second cell with additional content.");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Iterate through each row in the table and disable breaking across pages.
            foreach (Row row in table.Rows)
            {
                // Setting AllowBreakAcrossPages to false keeps the entire row together on a single page.
                row.RowFormat.AllowBreakAcrossPages = false;
            }

            // Save the document to the local file system.
            string outputPath = "PreventRowBreakAcrossPages.docx";
            doc.Save(outputPath);

            // Inform the user that the file has been created.
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
