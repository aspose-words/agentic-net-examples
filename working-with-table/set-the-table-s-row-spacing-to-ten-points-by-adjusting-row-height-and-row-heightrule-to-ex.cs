using System;
using System.IO;
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

            // Start a table and add the first row.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Row 1, Cell 1.");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2.");
            builder.EndRow();

            // Configure the second row to have a height of 10 points exactly.
            builder.RowFormat.Height = 10.0;
            builder.RowFormat.HeightRule = HeightRule.Exactly;

            // Add the second row.
            builder.InsertCell();
            builder.Write("Row 2, Cell 1.");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2.");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Verify that the second row has the expected height and rule.
            Row secondRow = table.Rows[1];
            if (Math.Abs(secondRow.RowFormat.Height - 10.0) > 0.001 ||
                secondRow.RowFormat.HeightRule != HeightRule.Exactly)
            {
                throw new InvalidOperationException("Row height or height rule was not set correctly.");
            }

            // Save the document.
            string outputPath = "TableRowSpacing.docx";
            doc.Save(outputPath);

            // Ensure the file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not saved.", outputPath);
        }
    }
}
