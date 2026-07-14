using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
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

            // ----- First row -----
            builder.InsertCell();
            builder.Write("First row, first cell.");
            builder.InsertCell();
            builder.Write("First row, second cell.");

            // End the first row and obtain the Row object.
            Row firstRow = builder.EndRow();

            // Set the height rule of this row to Auto (no explicit height is set).
            firstRow.RowFormat.HeightRule = HeightRule.Auto;

            // ----- Second row (optional) -----
            builder.InsertCell();
            builder.Write("Second row, first cell.");
            builder.InsertCell();
            builder.Write("Second row, second cell.");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document to a file.
            doc.Save("RowHeightAuto.docx");
        }
    }
}
