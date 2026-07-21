using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableHeaderRepeatExample
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

            // ----- Header row (repeated on each page) -----
            // Set the HeadingFormat flag so this row repeats on every page.
            builder.RowFormat.HeadingFormat = true;
            // Center the text inside the header cells.
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            // Define a reasonable width for the header cells.
            builder.CellFormat.Width = 100;

            // First header cell.
            builder.InsertCell();
            builder.Write("Header Column 1");
            // Second header cell.
            builder.InsertCell();
            builder.Write("Header Column 2");
            // Finish the header row.
            builder.EndRow();

            // ----- Data rows (regular rows) -----
            // Reset formatting for the regular rows.
            builder.RowFormat.HeadingFormat = false;
            builder.CellFormat.Width = 50;
            builder.ParagraphFormat.ClearFormatting();

            // Add enough rows to make the table span more than one page.
            for (int i = 1; i <= 50; i++)
            {
                // First cell of the data row.
                builder.InsertCell();
                builder.Write($"Row {i}, Column 1");
                // Second cell of the data row.
                builder.InsertCell();
                builder.Write($"Row {i}, Column 2");
                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Save the document to the local file system.
            const string outputPath = "TableWithRepeatingHeader.docx";
            doc.Save(outputPath);

            // Optional: verify that the file was created.
            if (System.IO.File.Exists(outputPath))
            {
                Console.WriteLine($"Document saved successfully to '{outputPath}'.");
            }
            else
            {
                throw new InvalidOperationException("Failed to save the document.");
            }
        }
    }
}
