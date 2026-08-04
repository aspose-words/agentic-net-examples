using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableRowSpacing
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

            // Set row formatting to achieve double line spacing.
            // HeightRule.Auto allows the row to expand if needed,
            // while Height = 24 points approximates double spacing (12 pt per line).
            builder.RowFormat.HeightRule = HeightRule.Auto;
            builder.RowFormat.Height = 24;

            // First row.
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            builder.EndRow();

            // Second row (inherits the same RowFormat settings).
            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Save the document.
            string outputPath = "TableRowSpacing.docx";
            doc.Save(outputPath);
        }
    }
}
