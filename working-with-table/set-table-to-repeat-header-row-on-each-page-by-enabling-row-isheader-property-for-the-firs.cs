using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableHeaderExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            Table table = builder.StartTable();

            // Mark the first row as a heading row that will repeat on each page.
            builder.RowFormat.HeadingFormat = true;

            // First header row.
            builder.InsertCell();
            builder.Write("Header Column 1");
            builder.InsertCell();
            builder.Write("Header Column 2");
            builder.EndRow();

            // Subsequent rows should not repeat as headings.
            builder.RowFormat.HeadingFormat = false;

            // Add enough rows to make the table span multiple pages.
            for (int i = 1; i <= 50; i++)
            {
                builder.InsertCell();
                builder.Write($"Row {i}, Column 1");
                builder.InsertCell();
                builder.Write($"Row {i}, Column 2");
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableHeaderRepeat.docx");
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"Document saved successfully to: {outputPath}");
            }
            else
            {
                throw new InvalidOperationException("Failed to save the document.");
            }
        }
    }
}
