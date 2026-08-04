using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableHeaderExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building a table.
            Table table = builder.StartTable();

            // ----- Header rows (will repeat on each page) -----
            // Enable the heading format for the rows that should repeat.
            builder.RowFormat.HeadingFormat = true;
            // Optional: center the text inside the header cells.
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            // Set a width for the header cells.
            builder.CellFormat.Width = 100;

            // First header cell.
            builder.InsertCell();
            builder.Write("Header Column 1");
            // End the first header row.
            builder.EndRow();

            // Second header cell.
            builder.InsertCell();
            builder.Write("Header Column 2");
            builder.EndRow();

            // ----- Normal rows (do not repeat) -----
            // Turn off the heading format for subsequent rows.
            builder.RowFormat.HeadingFormat = false;
            // Reset paragraph formatting to defaults.
            builder.ParagraphFormat.ClearFormatting();
            // Set a narrower width for regular cells.
            builder.CellFormat.Width = 50;

            // Add enough rows to make the table span multiple pages.
            for (int i = 1; i <= 30; i++)
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
            doc.Save("TableWithRepeatingHeader.docx");
        }
    }
}
