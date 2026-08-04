using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace RowSpacingExample
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

            // First row.
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.EndRow();

            // Third row.
            builder.InsertCell();
            builder.Write("Row 3, Cell 1");
            builder.InsertCell();
            builder.Write("Row 3, Cell 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // -----------------------------------------------------------------
            // NOTE:
            // Aspose.Words RowFormat does not expose SpaceBefore/SpaceAfter
            // properties. To influence the visual distance between rows you can
            // adjust the row height (Height) together with the HeightRule, or
            // insert empty rows as spacers. Below we demonstrate setting a
            // minimum height for each row to create extra space.
            // -----------------------------------------------------------------

            // Row 0: add extra space by increasing its height.
            table.Rows[0].RowFormat.Height = 30;               // height in points
            table.Rows[0].RowFormat.HeightRule = HeightRule.AtLeast;

            // Row 1: larger height for more spacing.
            table.Rows[1].RowFormat.Height = 40;
            table.Rows[1].RowFormat.HeightRule = HeightRule.AtLeast;

            // Row 2: even larger height.
            table.Rows[2].RowFormat.Height = 50;
            table.Rows[2].RowFormat.HeightRule = HeightRule.AtLeast;

            // Save the document.
            const string outputPath = "RowSpacing.docx";
            doc.Save(outputPath);
        }
    }
}
