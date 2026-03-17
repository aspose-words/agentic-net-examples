using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableRowSpacingExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new blank document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and add the first row (default formatting)
            builder.StartTable();
            builder.InsertCell();
            builder.Write("First row, cell 1.");
            builder.InsertCell();
            builder.Write("First row, cell 2.");
            builder.EndRow();

            // Configure the next row to have double line spacing.
            // HeightRule.Auto lets the row grow automatically, and Height is set to 24 points
            // (double the default 12‑point line height).
            builder.RowFormat.HeightRule = HeightRule.Auto;
            builder.RowFormat.Height = 24;

            // Add the second row with the new formatting
            builder.InsertCell();
            builder.Write("Second row, cell 1.");
            builder.InsertCell();
            builder.Write("Second row, cell 2.");
            builder.EndRow();

            // Finish the table
            builder.EndTable();

            // Save the document
            doc.Save("TableRowDoubleSpacing.docx");
        }
    }
}
